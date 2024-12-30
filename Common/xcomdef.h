#pragma once
#include <comdef.h>
#include <shobjidl.h>

// XXX: MSVC COM support lib does not provide interop for shobjidl so have to improvise

#define _XCOM_RAISE_IF_FAILED(expr) { \
    if(FAILED(hr = (expr))) \
        _com_raise_error(hr); \
}

#define _XCOM_RAISE_IF_WIN32(expr) { \
    if(expr) \
        _com_raise_error(HRESULT_FROM_WIN32(GetLastError())); \
}

_COM_SMARTPTR_TYPEDEF(IShellItem, __uuidof(IShellItem));
_COM_SMARTPTR_TYPEDEF(IFileOpenDialog, __uuidof(IFileOpenDialog));
_COM_SMARTPTR_TYPEDEF(IFileSaveDialog, __uuidof(IFileSaveDialog));
_COM_SMARTPTR_TYPEDEF(IFileDialogCustomize, __uuidof(IFileDialogCustomize));

