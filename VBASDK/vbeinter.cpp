#define ISOLATION_AWARE_ENABLED 1
#include "vbeinter.h"

#include <msi.h>

#pragma comment(lib, "msi.lib")

// ---

#ifdef VBA7
#   define VBA_COMPONENT_ID "{5999F473-6855-46D1-8093-B1C463B1DE9B}" // VBA71.msi, Component VBA_Vbe
#   define VBA_CATEGORY_ID "{D304E920-A68A-11D1-B5B5-006097C998E7}" // VBA71.msi, PublishComponent VBA_Vbe
#   define VBA_QUALIFIER "vbe.dll_7.1" // VBA71.msi, PublishComponent VBA_Vbe
#   define VBA_FILENAME "VBE7.DLL"
#   define VBA_REGNAME "Vbe71DllPath"
#else
#   define VBA_COMPONENT_ID "{CC29E953-7BC2-11D1-A921-00A0C91E2AA2}" // VBAOF11.msi, Component Global_VBA_Core
#   define VBA_CATEGORY_ID "{D304E920-A68A-11D1-B5B5-006097C998E7}" // VBAOF11.msi, PublishComponent Global_VBA_Core
#   define VBA_QUALIFIER "vbe.dll_6.0" // VBAOF11.msi, PublishComponent Global_VBA_Core
#   define VBA_FILENAME "VBE6.DLL"
#   define VBA_REGNAME "Vbe6DllPath"
#endif

#ifdef VBADEBUG
#   define VBA_DEBUG TRUE
#else
#   define VBA_DEBUG FALSE
#endif

// ---

typedef HRESULT(WINAPI* DLL_VBE_INIT)(DWORD uVerMaj, DWORD uVerMin, DWORD uVerRef, BOOL fDebug, VBE::VBACOOKIE dwCookie, VBE::VBAINITINFO* pVbaInitInfo, VBE::IVba** ppVba);
typedef HRESULT(WINAPI* DLL_VBA_INIT_HOST_ADDINS)(VBE::VBAINITHOSTADDININFO* pInitAddinInfo, VBE::IVbaHostAddins** ppHostAddins);
typedef HRESULT(WINAPI* DLL_VBE_TERM)(VBE::IVba* pVba);
typedef HRESULT(WINAPI* DLL_VBE_GET_HASH_OF_CODE)(VBE::VBACOOKIE dwCookie, IStorage* pStg, VBE::VBAHASH* pVbaHash);
typedef HRESULT(WINAPI* DLL_CAN_UNLOAD_NOW)();

static BOOL s_fVbaInitialized;
static HMODULE g_hInstVbe;

#ifdef VBASA

// Default load hook that does nothing
// https://devblogs.microsoft.com/oldnewthing/20200731-00/?p=104024
#pragma comment(linker, "/ALTERNATENAME:_VbaLoadHook=_DefaultVbaLoadHook")

extern "C"
bool _DefaultVbaLoadHook()
{
    return true;
}

#endif

static UINT _ReinstallProductFiles()
{
    CHAR szProductBuff[39];
    UINT ret = MsiGetProductCodeA(VBA_COMPONENT_ID, szProductBuff);
    if(!ret)
        return MsiReinstallFeatureA(szProductBuff, "ProductFiles", REINSTALLMODE_USERDATA | REINSTALLMODE_MACHINEDATA | REINSTALLMODE_FILEOLDERVERSION);
    return ret;
}

static UINT TCOQualifiedComponentState(LPCSTR sczGuid, LPCSTR sczQualifier, LPSTR szPath)
{
    DWORD cchPath = 260;
    UINT ret = MsiProvideQualifiedComponentA(sczGuid, sczQualifier, INSTALLMODE_NODETECTION, szPath, &cchPath);
    switch (ret)
    {
    case ERROR_SUCCESS:
        break;
    case ERROR_FILE_NOT_FOUND:
    case ERROR_INSTALL_FAILURE:
        cchPath = 260;
        if (!MsiProvideQualifiedComponentA(sczGuid, sczQualifier, INSTALLMODE_DEFAULT, szPath, &cchPath))
            break;
        return _ReinstallProductFiles();
    case ERROR_UNKNOWN_COMPONENT:
    case ERROR_INDEX_ABSENT:
        return ERROR_UNKNOWN_COMPONENT;
    default:
        return _ReinstallProductFiles();
    }

    OFSTRUCT ofStruct{ sizeof(OFSTRUCT) };
    if (OpenFile(szPath, &ofStruct, OF_EXIST) != HFILE_ERROR)
        return S_OK;

    cchPath = 260;
    ret = MsiProvideQualifiedComponentA(sczGuid, sczQualifier, INSTALLMODE_EXISTING, szPath, &cchPath);
    if (ret)
    {
        if (ret == ERROR_INSTALL_SOURCE_ABSENT)
            return _ReinstallProductFiles();
    }
    else if (OpenFile(szPath, &ofStruct, OF_EXIST) != HFILE_ERROR)
    {
        return S_OK;
    }

    cchPath = 260;
    if (!MsiProvideQualifiedComponentA(sczGuid, sczQualifier, REINSTALLMODE_USERDATA | REINSTALLMODE_MACHINEDATA | REINSTALLMODE_FILEOLDERVERSION, szPath, &cchPath) && OpenFile(szPath, &ofStruct, OF_EXIST) != HFILE_ERROR)
        return S_OK;
    return _ReinstallProductFiles();
}

static HMODULE DarwinLoadVbeDll()
{
    if (!g_hInstVbe)
    {
        CHAR psVBACompnPath[MAX_PATH];
        if(TCOQualifiedComponentState(VBA_CATEGORY_ID, VBA_QUALIFIER, psVBACompnPath) == ERROR_SUCCESS)  // VBAOF11.msi, PublishComponent Global_VBA_Core
            g_hInstVbe = LoadLibraryExA(psVBACompnPath, NULL, LOAD_WITH_ALTERED_SEARCH_PATH);
        if(!g_hInstVbe)
            g_hInstVbe = LoadLibraryA(VBA_FILENAME);
    }
    return g_hInstVbe;
}

static HMODULE LoadVbeDllFromRegistry()
{
    HRESULT hr = S_OK;
    HKEY hkVbeDllPath = NULL;
    if (RegOpenKeyA(HKEY_LOCAL_MACHINE, "Software\\Microsoft\\VBA", &hkVbeDllPath))
        hr = E_FAIL;

    CHAR szModulePath[MAX_PATH];
    DWORD cbSize = MAX_PATH;
    if (RegQueryValueExA(hkVbeDllPath, VBA_REGNAME, 0, 0, (LPBYTE)szModulePath, &cbSize))
        hr = E_FAIL;

    RegCloseKey(hkVbeDllPath);
    if (FAILED(hr))
    {
        return NULL;
    }
    else
    {
        return LoadLibraryExA(szModulePath, NULL, LOAD_WITH_ALTERED_SEARCH_PATH);
    }
}

static void LoadVbeDll()
{
    if (!g_hInstVbe)
    {
        g_hInstVbe = LoadVbeDllFromRegistry();
        if (!g_hInstVbe)
            g_hInstVbe = DarwinLoadVbeDll();
    }
}

STDAPI VbaInitialize(VBE::VBACOOKIE dwCookie, VBE::VBAINITINFO* pVbaInitInfo, VBE::IVba** ppVba)
{
    if (s_fVbaInitialized)
        return E_UNEXPECTED;
    if (!ppVba || !pVbaInitInfo || !dwCookie)
        return E_INVALIDARG;

    BOOL fVbeAlreadyLoaded = FALSE;
    *ppVba = NULL;
    if (!g_hInstVbe)
    {
        LoadVbeDll();
        if (!g_hInstVbe)
        {
            return HRESULT_FROM_WIN32(GetLastError());
        }

#ifdef VBASA
        if (!_VbaLoadHook())
        {
            HRESULT hr = HRESULT_FROM_WIN32(GetLastError());
            FreeLibrary(g_hInstVbe);
            g_hInstVbe = NULL;
            return hr;
        }
#endif
    }
    else
    {
        fVbeAlreadyLoaded = TRUE;
    }

    auto DllVbeInit = (DLL_VBE_INIT)GetProcAddress(g_hInstVbe, "DllVbeInit");
    HRESULT hr = E_FAIL;
    VBE::IVba* pVbaTemp = NULL;
    // XXX: The version code here seems not affecting much (vbeintr_s.obj uses 6.0.1032, APC71.DLL uses 6.0.1118)
    if (DllVbeInit && SUCCEEDED(hr = DllVbeInit(6, 0, 0, VBA_DEBUG, dwCookie, pVbaInitInfo, &pVbaTemp)))
    {
        *ppVba = pVbaTemp;
        s_fVbaInitialized = TRUE;
        return S_OK;
    }
    else
    {
        if (!fVbeAlreadyLoaded)
        {
            if (g_hInstVbe)
            {
                FreeLibrary(g_hInstVbe);
                g_hInstVbe = NULL;
            }
        }
        return hr;
    }
}

STDAPI VbaInitializeHostAddins(VBE::VBAINITHOSTADDININFO* pInitAddinInfo, VBE::IVbaHostAddins** ppHostAddins)
{
    if (!ppHostAddins)
        return E_FAIL;

    BOOL fVbeAlreadyLoaded = FALSE;
    *ppHostAddins = NULL;
    if (!g_hInstVbe)
    {
        LoadVbeDll();
        if (!g_hInstVbe)
        {
            return HRESULT_FROM_WIN32(GetLastError());
        }
    }
    else
    {
        fVbeAlreadyLoaded = TRUE;
    }

    auto DllVbaInitHostAddins = (DLL_VBA_INIT_HOST_ADDINS)GetProcAddress(g_hInstVbe, "DllVbaInitHostAddins");
    HRESULT hr = E_FAIL;
    if (DllVbaInitHostAddins && SUCCEEDED(hr = DllVbaInitHostAddins(pInitAddinInfo, ppHostAddins)))
    {
        return hr;
    }
    else
    {
        hr = E_FAIL;
    }
    if (!fVbeAlreadyLoaded)
    {
        if (g_hInstVbe)
        {
            FreeLibrary(g_hInstVbe);
            g_hInstVbe = NULL;
        }
    }
    return hr;
}

STDAPI VbaTerminate(VBE::IVba* pVba)
{
    if (!s_fVbaInitialized)
        return E_UNEXPECTED;
    if (!pVba)
        return E_INVALIDARG;
    if (!g_hInstVbe)
        return E_UNEXPECTED;

    auto DllVbeTerm = (DLL_VBE_TERM)GetProcAddress(g_hInstVbe, "DllVbeTerm");
    if (!DllVbeTerm)
        return E_UNEXPECTED;
    if (FAILED(DllVbeTerm(pVba)))
        return E_UNEXPECTED;
    
    auto DllCanUnloadNow = (DLL_CAN_UNLOAD_NOW)GetProcAddress(g_hInstVbe, "DllCanUnloadNow");
    if (!DllCanUnloadNow)
        return E_UNEXPECTED;
    if (DllCanUnloadNow() == S_OK)
    {
        FreeLibrary(g_hInstVbe);
        g_hInstVbe = NULL;
    }

    s_fVbaInitialized = FALSE;
    return S_OK;
}

STDAPI VbeGetHashOfCode(VBE::VBACOOKIE dwCookie, IStorage* pStg, VBE::VBAHASH* pVbaHash)
{
    LoadVbeDll();
    if (!g_hInstVbe)
    {
        return HRESULT_FROM_WIN32(GetLastError());
    }

    HRESULT hr = S_OK;
    auto DllVbeGetHashOfCode = (DLL_VBE_GET_HASH_OF_CODE)GetProcAddress(g_hInstVbe, "DllVbeGetHashOfCode");
    if (DllVbeGetHashOfCode)
        hr = DllVbeGetHashOfCode(dwCookie, pStg, pVbaHash);
    FreeLibrary(g_hInstVbe);
    g_hInstVbe = NULL; // XXX: Not in original code
    return hr;
}

#ifdef VBA7

STDAPI VbeSetGimmeTable()
{
    return S_OK;
}

#else

STDAPI VbeSetGimmeTable()
{
    HMODULE hInstMso9 = GetModuleHandleA("MSO9.DLL");
    if (!hInstMso9)
    {
        CHAR szPath[MAX_PATH];
        if (TCOQualifiedComponentState("{CC29E943-7BC2-11D1-A921-00A0C91E2AA2}", "mso.dll_9.0", szPath))
        {
            hInstMso9 = LoadLibraryA("MSO9.DLL");
            if (!hInstMso9)
                return E_FAIL;
        }
    }

    HRESULT hr = S_OK;
    auto MsoDwGetGimmeTableVersion = (void(WINAPI*)(DWORD))GetProcAddress(hInstMso9, "_MsoDwGetGimmeTableVersion@4");
    if (MsoDwGetGimmeTableVersion)
    {
        MsoDwGetGimmeTableVersion(0);
        hr = E_FAIL;
    }
    if (hInstMso9)
        FreeLibrary(hInstMso9);
    return hr;
}

#endif

