#include "stdafx.h"
#include "vbeauxmnu.h"

#include "iathook.h"

static HMODULE g_hmodVbe7Intl;
static void** g_pIatVbeLoadStringA;

static int WINAPI xLoadStringA(HINSTANCE hInstance, UINT uID, LPSTR lpBuffer, int cchBufferMax)
{
    if (!g_hmodVbe7Intl)
    {
        g_hmodVbe7Intl = GetModuleHandleW(L"VBE7INTL.DLL");
    }

    // Call site: VBE7!FSetLabelForControl -> VBE7!IntlLpstrStringOfID(20)
    // "Close & Return to %s(&C)"
    if (hInstance == g_hmodVbe7Intl && uID == 13000 + 20)
    {
        uID = 19160; // "Close(&C)"

        // VBE7!IntlLpstrStringOfID caches the string
        iat_modify(g_pIatVbeLoadStringA, LoadStringA);
    }

    return LoadStringA(hInstance, uID, lpBuffer, cchBufferMax);
}

bool init_vbeaux_mnu()
{
    auto hmodVbe7 = GetModuleHandleW(L"VBE7.DLL");
    if (!hmodVbe7)
    {
        return false;
    }
    
    g_pIatVbeLoadStringA = iat_get(hmodVbe7, "USER32.DLL", "LoadStringA");
    if (!g_pIatVbeLoadStringA)
    {
        return false;
    }

    if (!iat_modify(g_pIatVbeLoadStringA, xLoadStringA))
    {
        return false;
    }

    return true;
}

