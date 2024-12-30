#pragma once
// NOTE: Include this file after resource.h

#include <comdef.h>

extern HWND g_hwndMainWindow;

_bstr_t ui_load_string(UINT uiResId);

int ui_message_box_ex(UINT uType, UINT uiTextResId, UINT uiCaptionResId, ...);

// ---

#ifndef IDS_APPNAME
#   define __IDS_APPNAME_NOT_DEFINED
#   define IDS_APPNAME 0
#endif

#define ui_message_box(uType, uiTextResId, ...) \
    ui_message_box_ex((uType), (uiTextResId), (IDS_APPNAME), __VA_ARGS__)

#ifdef __IDS_APPNAME_NOT_DEFINED
#   undef __IDS_APPNAME_NOT_DEFINED
#   undef IDS_APPNAME
#endif

