#include "stdafx.h"
#include "resource.h"

#include "ui.h"
#include "vbeui.h"
#include "lickey.h"
#include "appvisv.h"
#include "vbeauxcmp.h"
#include "vbeauxmnu.h"
#include "vbeauxend.h"

#include "VbaHelper.h"
#include "CoInitScope.h"

#pragma comment(lib, "shlwapi.lib")

#pragma warning(disable: 28251) // Incompatible annotation on wWinMain

static _bstr_t get_config_path()
{
    WCHAR szPath[MAX_PATH];
    if (SearchPathW(NULL, L"VBASA.ini", NULL, _countof(szPath), szPath, NULL) == 0)
    {
        WCHAR szModulePath[MAX_PATH];
        if (GetModuleFileNameW(NULL, szModulePath, _countof(szModulePath)) == 0)
        {
            return L"VBASA.ini";
        }

        LPWSTR pszFilePart;
        if (GetFullPathNameW(szModulePath, _countof(szPath), szPath, &pszFilePart) == 0)
        {
            return L"VBASA.ini";
        }

        *pszFilePart = L'\0';
        wcscat_s(szPath, L"VBASA.ini");
    }
    return szPath;
}

int WINAPI wWinMain(HINSTANCE hInstance, HINSTANCE hPrevInstance, LPWSTR lpCmdLine, int nShowCmd)
{
    // NOTE: Command line from wWinMain() might not contain argv[0], thus not used
    int argc;
    WCHAR** argv = CommandLineToArgvW(GetCommandLineW(), &argc);
    if (!argv)
    {
        ui_message_box(MB_ICONERROR, IDS_ARGVFAIL, GetLastError());
        return 1;
    }

    bool noaux0 = false;
    bool noaux1 = false;
    _bstr_t filename;
    for (int i = 1; i < argc; i++)
    {
        if (argv[i][0] != L'/')
        {
            filename = argv[i];

            // TODO: Arguments?
            // Currently no way to implement; ExecutingMadeDllFrame() == 1 locks Command$() to vbNullString; see VBE7!rtcCommandBstr
            break;
        }
        else if (!_wcsicmp(argv[i], L"/?"))
        {
            ui_message_box(MB_ICONINFORMATION, IDS_HELPTEXT);
            return 0;
        }
        else if (!_wcsicmp(argv[i], L"/f0"))
        {
            noaux0 = true;
        }
        else if (!_wcsicmp(argv[i], L"/f1"))
        {
            noaux1 = true;
        }
        else
        {
            ui_message_box(MB_ICONERROR, IDS_ARGFAIL, argv[i]);
            return 1;
        }
    }

    LocalFree(argv);

    // ---

    _bstr_t config = get_config_path();

    WCHAR szLicKey[MAX_PATH];
    GetPrivateProfileStringW(L"VBASA", L"LicenseKey", L"", szLicKey, _countof(szLicKey), config);
    if (!init_lickey(szLicKey))
    {
        return 1;
    }
    WritePrivateProfileStringW(L"VBASA", L"LicenseKey", g_bstrLicKey, config);

    // ---

    int ret = 0;
    CoInitScope __coInitScope(COINIT_APARTMENTTHREADED);
    try {
        VbaHelperParams params;
        params.cookie = g_bstrLicKey;
        params.nameid = IDS_BLDAPPNAME;
        params.headless = false;

        try
        {
            VbaHelper::CreateInstance(params);
        }
        catch (const _com_error& exc)
        {
            if (init_appvisv())
            {
                try
                {
                    VbaHelper::CreateInstance(params);
                }
                catch (const _com_error& exc)
                {
                    ui_message_box(MB_ICONERROR, IDS_VBAFAIL, GetLastError(), exc.Error());
                }
            }
            else
            {
                ui_message_box(MB_ICONERROR, IDS_VBAFAIL, GetLastError(), exc.Error());
            }
        }

        if (!VbaHelper::GetInstance().GetVba())
        {
            if (ui_message_box(MB_ICONINFORMATION | MB_YESNO | MB_DEFBUTTON2, IDS_VBAHINT) == IDYES)
            {
                ShellExecuteW(NULL, L"open", ui_load_string(IDS_VBAURL), NULL, NULL, SW_NORMAL);
            }
            return 1;
        }

        if (!noaux0)
        {
            // CMProperties fix
            if (!init_vbeaux_cmp())
            {
                ui_message_box(MB_ICONWARNING, IDS_VBEAUXFAIL, GetLastError());
            }

            // Menu text fix
            if (!init_vbeaux_mnu())
            {
                ui_message_box(MB_ICONWARNING, IDS_VBEAUXFAIL, GetLastError());
            }

            if (!noaux1)
            {
                // End statement fix
                if (!init_vbeaux_end())
                {
                    ui_message_box(MB_ICONWARNING, IDS_VBEAUXFAIL, GetLastError());
                }
            }
        }

        // ---

        VBIDE::VBEPtr vbe(VbaHelper::GetInstance().GetVba()->GetExtensibilityObject());
        if (!init_vbeui(vbe))
        {
            ui_message_box(MB_ICONWARNING, IDS_VBEUIFAIL, GetLastError());
        }

        if (!!filename)
        {
            try
            {
                vbe->VBProjects->Open(filename);
            }
            catch (const _com_error& exc)
            {
                ui_message_box(MB_ICONERROR, IDS_LOADFAIL, exc.Error(), (wchar_t*)exc.Description());
            }
        }
        else
        {
            vbe_new_project(vbe);
        }

        // Show VBE
        VBE::IUIElementPtr ui;
        IServiceProviderPtr(VbaHelper::GetInstance().GetVba())->QueryService(SID_SVbaMainUI, &ui);
        ui->Show();

        ret = (int)VbaHelper::GetInstance().MessageLoop(); // HACK: See VbaCompManagerSite::raw_PushMessageLoop
    }
    catch (const _com_error& exc)
    {
        ui_message_box(MB_ICONERROR, IDS_COMERROR, exc.Error(), (wchar_t*)exc.Description());
    }

    return ret;
}

