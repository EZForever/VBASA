#include "stdafx.h"
#include "VbaSite.h"

#pragma comment(lib, "htmlhelp.lib")

#if defined(_M_IX86)
#   define xFOLDERID_ProgramFilesCommonX86 FOLDERID_ProgramFilesCommon
#elif defined(_M_AMD64)
#   define xFOLDERID_ProgramFilesCommonX86 FOLDERID_ProgramFilesCommonX86
#else
#   error "Unsupported platform"
#endif

static wchar_t g_szVba71X86[MAX_PATH];

VBE::IVbaSite* VbaSite::CreateInstance(bool headless, LCID lcid)
{
    auto ptr = new VbaSite(headless, lcid);
    ptr->AddRef();
    return ptr;
}

// --- Implements VbaSite->IUnknown

STDMETHODIMP VbaSite::QueryInterface(REFIID riid, void** ppvObject)
{
    if (!ppvObject)
        return E_POINTER;

    if (riid != IID_IUnknown && riid != __uuidof(VBE::IVbaSite))
    {
        *ppvObject = nullptr;
        return E_NOINTERFACE;
    }

    this->AddRef();
    *ppvObject = this;
    return S_OK;
}

STDMETHODIMP_(ULONG) VbaSite::AddRef()
{
    return _InterlockedIncrement(&m_refcnt);
}

STDMETHODIMP_(ULONG) VbaSite::Release()
{
    LONG lref = _InterlockedDecrement(&m_refcnt);
    if (lref == 0)
        delete this;
    return lref;
}

// --- Implements VbaSite

STDMETHODIMP VbaSite::raw_GetOwnerWindow(wireHWND* phwnd)
{
    *phwnd = NULL;
    return S_OK;
}

STDMETHODIMP VbaSite::raw_Notify(VBE::VBASITENOTIFY vbahnotify)
{
    return S_OK;
}

STDMETHODIMP VbaSite::raw_FindFile(LPWSTR lpstrFileName, BSTR* pbstrFullPath)
{
    return E_NOTIMPL;
}

STDMETHODIMP VbaSite::raw_GetMiniBitmap(GUID* rgguid, wireHBITMAP* phbmp, unsigned long* pcrMask)
{
    return E_NOTIMPL;
}

STDMETHODIMP VbaSite::raw_CanEnterBreakMode(VBE::VBA_BRK_INFO* pbrkinfo)
{
    if (m_headless)
        pbrkinfo->brkopt = VBE::VBA_BRKOPT_End;
    return S_OK;
}

STDMETHODIMP VbaSite::raw_SetStepMode(VBE::VBASTEPMODE vbastepmode)
{
    return S_OK;
}

STDMETHODIMP VbaSite::raw_ShowHide(long fVisible)
{
    if (!m_headless && !fVisible)
        PostQuitMessage(0);
    return S_OK;
}

STDMETHODIMP VbaSite::raw_InstanceCreated(IUnknown* punkInstance)
{
    return S_OK;
}

STDMETHODIMP VbaSite::raw_HostCheckReference(long fSave, GUID* rgguid, UINT* puVerMaj, UINT* puVerMin)
{
    return S_OK;
}

STDMETHODIMP VbaSite::raw_GetIEVersion(LONG* plMajVer, LONG* plMinVer, long* pfCanInstall)
{
    return E_NOTIMPL;
}

STDMETHODIMP VbaSite::raw_UseIEFeature(LONG lMajVer, LONG lMinVer, wireHWND hwndOwner)
{
    return E_NOTIMPL;
}

STDMETHODIMP VbaSite::raw_ShowHelp(wireHWND hwnd, LPWSTR szHelp, UINT wCmd, DWORD dwHelpContext, long fWinHelp)
{
    // XXX: Due to oversights in both VBE and AppvIsv, help files are searched for only in ProgramFilesCommonX64 on a x64 system
    // Where VBE help files are located in ProgramFilesCommonX86
    if (g_szVba71X86[0] == L'\0')
    {
        PWSTR pszCommonFilesX86 = NULL;
        auto hr = SHGetKnownFolderPath(xFOLDERID_ProgramFilesCommonX86, 0, NULL, &pszCommonFilesX86);
        if (FAILED(hr))
        {
            // "The calling process is responsible for freeing [...] by calling CoTaskMemFree, whether SHGetKnownFolderPath succeeds or not."
            if (pszCommonFilesX86)
                CoTaskMemFree(pszCommonFilesX86);
            return hr;
        }

        wcscpy_s(g_szVba71X86, pszCommonFilesX86);
        wcscat_s(g_szVba71X86, L"\\Microsoft Shared\\VBA\\VBA7.1");
        CoTaskMemFree(pszCommonFilesX86);
    }

    wchar_t szHelpFileName[MAX_PATH];
    swprintf_s(szHelpFileName, L"%s\\%u\\%s", g_szVba71X86, m_lcid, szHelp);

    // HelpFileName may contain window or topic names
    wchar_t szFileName[MAX_PATH];
    wcscpy_s(szFileName, szHelpFileName);
    szFileName[_countof(szFileName) - 1] = L'\0'; // Make IntelliSense happy

    wchar_t* pszFileNamePart;
    pszFileNamePart = wcsstr(szFileName, L"::");
    if (pszFileNamePart)
        *pszFileNamePart = L'\0';
    pszFileNamePart = wcsstr(szFileName, L"[>");
    if (pszFileNamePart)
        *pszFileNamePart = L'\0';

    if ((GetFileAttributesW(szFileName) & (FILE_ATTRIBUTE_DIRECTORY | FILE_ATTRIBUTE_DEVICE)) != 0)
    {
        // Help file is not there; do nothing and let APC take care of it
        return S_OK;
    }

    if (fWinHelp)
    {
        if (!WinHelpW((HWND)hwnd, szHelpFileName, wCmd, dwHelpContext))
            return E_FAIL;
    }
    else
    {
        if (!HtmlHelpW((HWND)hwnd, szHelpFileName, wCmd, dwHelpContext))
            return E_FAIL;
    }
    return S_OK;
}

STDMETHODIMP VbaSite::raw_OpenProject(wireHWND hwndOwner, VBE::IVbaProject* pVbaRefProj, LPWSTR lpstrFileName, VBE::IVbaProject** ppVbaProj)
{
    return E_NOTIMPL;
}

