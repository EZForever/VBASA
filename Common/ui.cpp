#include "stdafx.h"
#include "ui.h"

#include <varargs.h>

HWND g_hwndMainWindow;

// NOTE: On why sometimes valid resource strings cannot be found:
// Given facts:
// - String Tables are packed into pages, 16 strings each
// - Standard resource search order priorities Language Neutual resources
// - LoadString locates String Table page first, then search for the string in that page
// So if a string is numerically "in the same page" as another LN string, it might not be picked up by LoadString
_bstr_t ui_load_string(UINT uiResId)
{
    LPWSTR pString = NULL;
    int cbString = LoadStringW(GetModuleHandleW(NULL), uiResId, (LPWSTR)&pString, 0);
    if (cbString != 0)
    {
        // "String resources are not guaranteed to be null-terminated in the module's resource table."
        return _bstr_t(SysAllocStringLen(pString, cbString), false);
    }
    else
    {
        WCHAR szString[4096] = { 0 };
        DWORD_PTR args[] = { uiResId, GetLastError() };
        FormatMessageW(FORMAT_MESSAGE_FROM_STRING | FORMAT_MESSAGE_ARGUMENT_ARRAY, L"[MissingString#%1!u!,%2!u!]", 0, 0, szString, _countof(szString), (va_list*)args);
        return _bstr_t(szString);
    }
}

int ui_message_box_ex(UINT uType, UINT uiTextResId, UINT uiCaptionResId, ...)
{
    va_list args;
    va_start(args, uiCaptionResId);

    WCHAR szString[4096] = { 0 };
    FormatMessageW(FORMAT_MESSAGE_FROM_STRING, (LPWSTR)ui_load_string(uiTextResId), 0, 0, szString, _countof(szString), &args);

    va_end(args);

    return MessageBoxW(g_hwndMainWindow, szString, ui_load_string(uiCaptionResId), uType);
}

