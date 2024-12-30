#include "stdafx.h"
#include "vbeauxcmp.h"

#include "iathook.h"

#include "FauxTypeLib2.h"

static void** g_pIatVbeLoadTypeLib;

static HRESULT STDAPICALLTYPE xLoadTypeLib(LPCOLESTR szFile, ITypeLib** pptlib)
{
    auto hr = LoadTypeLib(szFile, pptlib);
    if (SUCCEEDED(hr))
    {
        wchar_t szFileUpper[MAX_PATH];
        wcscpy_s(szFileUpper, szFile);
        szFileUpper[_countof(szFileUpper) - 1] = L'\0'; // Make IntelliSense happy
        _wcsupr_s(szFileUpper);

        wchar_t* pszTarget = wcsstr(szFileUpper, L"\\VBE7.DLL\\3");
        if (pszTarget && !wcscmp(pszTarget, L"\\VBE7.DLL\\3"))
        {
            auto tlib = *pptlib;
            try
            {
                *pptlib = FauxTypeLib2::CreateInstance(tlib);
            }
            catch (const _com_error&)
            {
                return hr;
            }
            tlib->Release();

            // Seems that this typelib gets cached by VBE7 (or VBEUI?), no need to keep this hook
            iat_modify(g_pIatVbeLoadTypeLib, LoadTypeLib);
        }
    }
    return hr;
}

bool init_vbeaux_cmp()
{
    auto hmodVbe7 = GetModuleHandleW(L"VBE7.DLL");
    if (!hmodVbe7)
    {
        return false;
    }

    g_pIatVbeLoadTypeLib = iat_get(hmodVbe7, "OLEAUT32.DLL", "LoadTypeLib");
    if (!g_pIatVbeLoadTypeLib)
    {
        g_pIatVbeLoadTypeLib = iat_get_ordinal(hmodVbe7, "OLEAUT32.DLL", 161);
        if (!g_pIatVbeLoadTypeLib)
        {
            return false;
        }
    }

    if (!iat_modify(g_pIatVbeLoadTypeLib, xLoadTypeLib))
    {
        return false;
    }

    return true;
}

