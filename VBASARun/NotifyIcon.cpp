#include "stdafx.h"
#include "resource.h"
#include "NotifyIcon.h"

#include "ui.h"
#include "xcomdef.h"

#include "VbaHelper.h"

ATOM NotifyIcon::g_atomMessageWindow;

LRESULT CALLBACK NotifyIcon::WndProc(HWND hWnd, UINT Msg, WPARAM wParam, LPARAM lParam)
{
	switch (Msg)
	{
	case WM_CREATE:
	{
		auto pThis = (NotifyIcon*)((LPCREATESTRUCTW)lParam)->lpCreateParams;
		SetWindowLongPtrW(hWnd, GWLP_USERDATA, (LONG_PTR)pThis);

		pThis->m_icondata.hWnd = hWnd;
		Shell_NotifyIconW(NIM_ADD, &pThis->m_icondata);
		Shell_NotifyIconW(NIM_SETVERSION, &pThis->m_icondata);

		return TRUE;
	}
	case WM_DESTROY:
	{
		auto pThis = (NotifyIcon*)GetWindowLongPtrW(hWnd, GWLP_USERDATA);

		Shell_NotifyIconW(NIM_DELETE, &pThis->m_icondata);

		return TRUE;
	}
	case WM_USER + 5432:
	{
		if (LOWORD(lParam) == WM_RBUTTONUP)
		{
			auto pThis = (NotifyIcon*)GetWindowLongPtrW(hWnd, GWLP_USERDATA);
			auto hmenu = LoadMenuW(NULL, MAKEINTRESOURCEW(IDR_ICONMENU));
			if (hmenu)
			{
				HMENU hmenuPopup = GetSubMenu(hmenu, 0);
#if 1
				MENUINFO menuInfo{ sizeof(menuInfo) };
				menuInfo.fMask = MIM_STYLE;
				menuInfo.dwStyle = MNS_NOCHECK;
				SetMenuInfo(hmenuPopup, &menuInfo);
#endif
				MENUITEMINFOW menuItemInfo{ sizeof(menuItemInfo) };
				menuItemInfo.fMask = MIIM_STRING;
				menuItemInfo.dwTypeData = pThis->m_icondata.szTip;
				SetMenuItemInfoW(hmenuPopup, 0, TRUE, &menuItemInfo);

				// See https://learn.microsoft.com/en-us/windows/win32/api/winuser/nf-winuser-trackpopupmenu#remarks
				SetForegroundWindow(hWnd);
				TrackPopupMenu(hmenuPopup, TPM_LEFTALIGN, GET_X_LPARAM(wParam), GET_Y_LPARAM(wParam), 0, hWnd, NULL);
				PostMessageW(hWnd, WM_NULL, 0, 0);

				DestroyMenu(hmenu);
			}
		}
		return TRUE;
	}
	case WM_COMMAND:
	{
		if (LOWORD(wParam) == ID_POPUP_FORCESTOP)
		{
			VbaHelper::GetInstance().Reset();
			return TRUE;
		}
		break;
	}
	default:
		break;
	}
	return DefWindowProcW(hWnd, Msg, wParam, lParam);
}

NotifyIcon::NotifyIcon(_bstr_t bstrProjectName)
{
	HINSTANCE hInstance = GetModuleHandleW(NULL);

	if (g_atomMessageWindow == 0)
	{
		WNDCLASSEXW cls{ sizeof(cls) };
		cls.lpszClassName = L"NotifyIcon";
		cls.hInstance = hInstance;
		cls.lpfnWndProc = WndProc;

		g_atomMessageWindow = RegisterClassExW(&cls);
		_XCOM_RAISE_IF_WIN32(g_atomMessageWindow == 0);
	}

	m_icondata = NOTIFYICONDATAW{ sizeof(m_icondata) };
	m_icondata.uVersion = NOTIFYICON_VERSION_4;
	m_icondata.uFlags = NIF_MESSAGE | NIF_ICON | NIF_TIP | NIF_SHOWTIP;
	m_icondata.uCallbackMessage = WM_USER + 5432;
	m_icondata.hIcon = LoadIconW(hInstance, MAKEINTRESOURCEW(IDI_APPICON));

	WCHAR szProjectName[MAX_PATH];
	if (bstrProjectName.length() == 0)
	{
		_XCOM_RAISE_IF_WIN32(GetModuleFileNameW(hInstance, szProjectName, _countof(szProjectName)) == 0);
	}
	else
	{
		wcscpy_s(szProjectName, bstrProjectName);
	}
	szProjectName[_countof(szProjectName) - 1] = L'\0'; // Make Intellisense happy

	WCHAR szProjectFullName[MAX_PATH];
	LPWSTR pszFilePart;
	_XCOM_RAISE_IF_WIN32(GetFullPathNameW(szProjectName, _countof(szProjectFullName), szProjectFullName, &pszFilePart) == 0);

	DWORD_PTR args[] = { (DWORD_PTR)pszFilePart };
	_XCOM_RAISE_IF_WIN32(FormatMessageW(FORMAT_MESSAGE_FROM_STRING | FORMAT_MESSAGE_ARGUMENT_ARRAY, (LPCWSTR)ui_load_string(IDS_ICONTITLE), 0, 0, m_icondata.szTip, _countof(m_icondata.szTip), (va_list*)args) == 0);

	m_hwnd = CreateWindowExW(0, MAKEINTATOM(g_atomMessageWindow), NULL, 0, 0, 0, 0, 0, HWND_MESSAGE, NULL, hInstance, this);
	_XCOM_RAISE_IF_WIN32(!m_hwnd);
}

NotifyIcon::~NotifyIcon()
{
	if (m_hwnd)
	{
		DestroyWindow(m_hwnd);
		m_hwnd = NULL;
	}
}

