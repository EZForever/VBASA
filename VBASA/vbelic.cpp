#include "stdafx.h"

#include "iathook.h"
#include "vbeinter.h"

// From Microsoft Office 2000 Developer CD
static const char g_szVbeLicKey[] = "8804558B-B773-11d1-BC3E-0000F87552E7";
static const char g_szVbeLicValue[] = "yjwqljkjjjljjjnjwklqqjqjxkrpljlqplvk";

static void** g_pIatVbeRegQueryValueA;

static LSTATUS APIENTRY xRegQueryValueA(HKEY hKey, LPCSTR lpSubKey, LPSTR lpData, PLONG lpcbData)
{
    // Grant VBA standalone license; see VBE7!VBEValidateLicense
    if (lpSubKey && !strcmp(lpSubKey, g_szVbeLicKey))
    {
        LSTATUS ret = ERROR_SUCCESS;
        LONG cbData = lpcbData ? *lpcbData : 0;
        if (cbData < sizeof(g_szVbeLicValue))
        {
            ret = ERROR_MORE_DATA;
        }

        // XXX: Doc did not specify what happens if lpcbData is NULL
        if (lpcbData)
        {
            *lpcbData = sizeof(g_szVbeLicValue);
            if (lpData)
            {
                strcpy_s(lpData, cbData, g_szVbeLicValue);

                iat_modify(g_pIatVbeRegQueryValueA, RegQueryValueA);
            }
        }
        return ret;
    }

    return RegQueryValueA(hKey, lpSubKey, lpData, lpcbData);
}

static bool vbe_lic_installed()
{
    HKEY hkLicenses;
    if (RegOpenKeyA(HKEY_CLASSES_ROOT, "Licenses", &hkLicenses) != ERROR_SUCCESS)
    {
        return false;
    }

    char szLicValue[_countof(g_szVbeLicValue)];
    LONG cbLicValue = sizeof(szLicValue);
    if (RegQueryValueA(hkLicenses, g_szVbeLicKey, szLicValue, &cbLicValue) != ERROR_SUCCESS)
    {
        RegCloseKey(hkLicenses);
        return false;
    }
    RegCloseKey(hkLicenses);

    return cbLicValue == sizeof(szLicValue) && !strcmp(szLicValue, g_szVbeLicValue);
}

// Hooks into VbaInitialize to install IAT hook ASAP
// See comment in vbeinter.h
extern "C"
bool _VbaLoadHook()
{
    if (vbe_lic_installed())
    {
        // License is there so do not bother reinstalling
        return true;
    }

    // Check if the license has already been installed
    auto hModule = GetModuleHandleW(L"VBE7.DLL");
    if (!hModule)
    {
        return false;
    }

    g_pIatVbeRegQueryValueA = iat_get(hModule, "ADVAPI32.DLL", "RegQueryValueA");
    if (!g_pIatVbeRegQueryValueA)
    {
        return false;
    }

    if (!iat_modify(g_pIatVbeRegQueryValueA, xRegQueryValueA))
    {
        return false;
    }

    return true;
}

