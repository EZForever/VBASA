#include "stdafx.h"
#include "resource.h"
#include "lickey.h"

#include "ui.h"
#include "evalkey.h"

// Enable additional, more pedantic checks than VBA6MTRT!VerifyCookie
//#define LICKEY_ADDITIONAL_CHECKS

_bstr_t g_bstrLicKey;

static INT_PTR CALLBACK InputLicDlgProc(HWND hWnd, UINT Msg, WPARAM wParam, LPARAM lParam)
{
    switch (Msg)
    {
    case WM_INITDIALOG:
    {
        g_hwndMainWindow = hWnd;

        int hres = GetSystemMetrics(SM_CXFULLSCREEN);
        int vres = GetSystemMetrics(SM_CYFULLSCREEN);
        if (hres != 0 && vres != 0)
        {
            RECT rc{ 0 };
            GetWindowRect(hWnd, &rc);

            int xpos = (hres - (rc.right - rc.left)) / 2;
            int ypos = (vres - (rc.bottom - rc.top)) / 3;
            SetWindowPos(hWnd, NULL, xpos, ypos, 0, 0, SWP_NOSIZE | SWP_NOZORDER);
        }

        SendDlgItemMessageW(hWnd, IDC_INPUTLIC_EDIT, EM_SETLIMITTEXT, MAX_PATH - 1, 0);
        SetFocus(GetDlgItem(hWnd, IDC_INPUTLIC_EDIT));

        MessageBeep(MB_ICONQUESTION);
        return TRUE;
    }
    case WM_DESTROY:
    {
        g_hwndMainWindow = NULL;
        return FALSE; // XXX: Let dialog manager do the rest of work
    }
    case WM_COMMAND:
    {
        switch (LOWORD(wParam))
        {
        case IDOK:
        {
            WCHAR szResult[MAX_PATH];
            UINT cbResult = GetDlgItemTextW(hWnd, IDC_INPUTLIC_EDIT, szResult, _countof(szResult));
            if (cbResult == 0)
            {
                if (ui_message_box(MB_ICONWARNING | MB_YESNO, IDS_LICEVALWARN) == IDYES)
                {
                    EndDialog(hWnd, (INT_PTR)SysAllocString(g_wszVbaEvalLicKey));
                }
                else
                {
                    SetFocus(GetDlgItem(hWnd, IDC_INPUTLIC_EDIT));
                }
            }
            else
            {
                EndDialog(hWnd, (INT_PTR)SysAllocString(szResult));
            }
            return TRUE;
        }
        case IDCANCEL:
        {
            EndDialog(hWnd, NULL);
            return TRUE;
        }
        default:
            break;
        }
    }
    default:
        break;
    }
    return FALSE;
}

static bool lickey_valid(_bstr_t lickey)
{
    auto len = lickey.length();
    if (len < 48 || len % 2 != 0)
    {
        return false;
    }
#ifdef LICKEY_ADDITIONAL_CHECKS
    if (len > 54)
    {
        return false;
    }
#endif

    auto str = lickey.GetBSTR();
    for (UINT i = 0; i < len; i++)
    {
        if (!iswxdigit(str[i]))
        {
            return false;
        }
    }

    BYTE s, ve, pke;
    if (swscanf_s(str, L"%02hhx%02hhx%02hhx", &s, &ve, &pke) != 3)
    {
        return false;
    }

    BYTE v = ve ^ s;
    BYTE pk = pke ^ s;
    UINT il = (s & 6) / 2;
    if (v != 1 || pk != 0x47)
    {
        return false;
    }
#ifdef LICKEY_ADDITIONAL_CHECKS
    if (il != (len - 48) / 2)
    {
        return false;
    }
#endif

    BYTE ub1 = pk;
    BYTE eb1 = pke;
    BYTE eb2 = ve;
    BYTE buf[3 + 4 + 17] = { 0 };
    for (UINT i = 0; i < il + 4 + 17; i++)
    {
#ifndef LICKEY_ADDITIONAL_CHECKS
        // Skip ignored additional byte
        if (i == il + 4 + 17 - 1)
        {
            break;
        }
#endif

        BYTE be;
        if (swscanf_s(str + (3 + i) * 2, L"%02hhx", &be) != 1)
        {
            return false;
        }

        BYTE b = be ^ (eb2 + ub1);
        buf[i] = b;

        eb2 = eb1;
        eb1 = be;
        ub1 = b;
    }

#ifdef LICKEY_ADDITIONAL_CHECKS
    for (UINT i = 0; i < il; i++)
    {
        if (buf[i] != s)
        {
            return false;
        }
    }
#endif

    DWORD dl = *(PDWORD)(buf + il);
#ifdef LICKEY_ADDITIONAL_CHECKS
    if (dl != 17)
    {
        return false;
    }
#else
    if (dl < 16)
    {
        return false;
    }
#endif

    PCHAR data = (PCHAR)(buf + il + 4);
    UINT y, m, d, n;
    DWORD e;
    if (sscanf_s(data, "%02u%02u%02u%02u%08x", &y, &m, &d, &n, &e) != 5)
    {
        return false;
    }
    if (m < 1 || m > 12 || d < 1 || d > 31 || n < 1)
    {
        return false;
    }

    tm lictm = { 0 };
    lictm.tm_year = y + (y >= 90) ? 0 : 100;
    lictm.tm_mon = m - 1;
    lictm.tm_mday = d;

    time_t currtm = time(NULL);
    if (difftime(mktime(&lictm), currtm) > 0)
    {
        return false;
    }
    if (difftime(e, currtm) < 0)
    {
        return false;
    }

#ifdef LICKEY_ADDITIONAL_CHECKS
    BYTE x = data[16];
    if (__popcnt(x) > 2)
    {
        return false;
    }
#endif

    return true;
}

bool init_lickey(_bstr_t lickey)
{
    while (true)
    {
        if (lickey_valid(lickey))
        {
            break;
        }

        if (lickey.length() != 0)
        {
            ui_message_box(MB_ICONWARNING, IDS_LICINVALID);
        }

        auto newlickey = (BSTR)DialogBoxW(GetModuleHandleW(NULL), MAKEINTRESOURCEW(IDD_INPUTLIC), NULL, InputLicDlgProc);
        lickey.Assign(newlickey);
        if (!newlickey)
        {
            break;
        }
    }

    g_bstrLicKey = lickey;
    return lickey.length() != 0;
}

