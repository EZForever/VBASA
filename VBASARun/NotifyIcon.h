#pragma once
#include "stdafx.h"

class NotifyIcon final
{
private:
	static ATOM g_atomMessageWindow;

	static LRESULT CALLBACK WndProc(HWND hWnd, UINT Msg, WPARAM wParam, LPARAM lParam);

	// ---

	NOTIFYICONDATAW m_icondata;
	HWND m_hwnd;

public:
	NotifyIcon(_bstr_t bstrProjectName);
	NotifyIcon(NotifyIcon&&) = delete;
	~NotifyIcon();
};

