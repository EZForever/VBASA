#include "stdafx.h"
#include "resource.h"

#include "ui.h"
#include "appvisv.h"
#include "xcomdef.h"
#include "compression.h"

#include "VbaHelper.h"
#include "NotifyIcon.h"
#include "CoInitScope.h"

#pragma comment(lib, "shlwapi.lib")

#pragma warning(disable: 28251) // Incompatible annotation on wWinMain

static buffer_t get_resource_project()
{
    HRSRC hrsrc = FindResourceW(NULL, L"VBAPROJ", RT_RCDATA);
    if (!hrsrc)
    {
        return buffer_t();
    }

    HGLOBAL hResData = LoadResource(NULL, hrsrc);
    if (!hResData)
    {
        return buffer_t();
    }

    auto pData = (PBYTE)LockResource(hResData);
    return buffer_t(pData, pData + SizeofResource(NULL, hrsrc));
}

static _bstr_t get_disk_project()
{
    HRESULT hr;

    auto bstrOpenProj = ui_load_string(IDS_OPENPROJ);
    auto bstrOpenAny = ui_load_string(IDS_OPENANY);
    COMDLG_FILTERSPEC rgFilterSpecs[] = {
        { bstrOpenProj.GetBSTR(), L"*.vba" },
        { bstrOpenAny.GetBSTR(), L"*.*" }
    };

    IFileOpenDialogPtr dlg(CLSID_FileOpenDialog);
    _XCOM_RAISE_IF_FAILED(dlg->SetTitle(ui_load_string(IDS_OPENTITLE)));
    _XCOM_RAISE_IF_FAILED(dlg->SetFileTypes(_countof(rgFilterSpecs), rgFilterSpecs));

    if (FAILED(hr = dlg->Show(NULL)))
    {
        // User cancelled the dialog. Abort
        return _bstr_t();
    }

    IShellItemPtr item;
    _XCOM_RAISE_IF_FAILED(dlg->GetResult(&item));

    LPWSTR pszName;
    _bstr_t result;
    _XCOM_RAISE_IF_FAILED(item->GetDisplayName(SIGDN_FILESYSPATH, &pszName));
    result = pszName;
    CoTaskMemFree(pszName);
    return result;
}

int WINAPI wWinMain(HINSTANCE hInstance, HINSTANCE hPrevInstance, LPWSTR lpCmdLine, int nShowCmd)
{
    _bstr_t filename;
    auto buf = get_resource_project();
    if (buf.empty())
    {
        // This program is directly run; check if a file is given on commandline
        
        // NOTE: Command line from wWinMain() might not contain argv[0], thus not used
        int argc;
        WCHAR** argv = CommandLineToArgvW(GetCommandLineW(), &argc);
        if (!argv)
        {
            ui_message_box(MB_ICONERROR, IDS_ARGVFAIL, GetLastError());
            return 1;
        }

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
            else
            {
                ui_message_box(MB_ICONERROR, IDS_ARGFAIL, argv[i]);
                return 1;
            }
        }

        LocalFree(argv);
    }

    int ret = 0;
    HRESULT hr;
    CoInitScope __coInitScope(COINIT_APARTMENTTHREADED);
    try
    {
        // If still no file was specified, show a Open dialog
        if (buf.empty() && filename.length() == 0)
        {
            filename = get_disk_project();

            // If STILL no file was specified, silently quit
            if (filename.length() == 0)
            {
                return 1;
            }
        }

        VbaHelperParams params;
        params.cookie = VBE::VBACOOKIE_OFFICE;
        params.nameid = IDS_APPNAME;
        params.headless = true;
        params.mdehack = !buf.empty();

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

        IStoragePtr stg;
        if (buf.empty())
        {
            if (FAILED(hr = StgOpenStorage(filename, NULL, STGM_READ | STGM_SHARE_DENY_WRITE, NULL, 0, &stg)))
            {
                ui_message_box(MB_ICONERROR, IDS_LOADFAIL1, hr);
                return 1;
            }
        }
        else
        {
            buf = decompress_buffer(buf);

            ILockBytesPtr lkbyt;
            if (FAILED(hr = CreateILockBytesOnHGlobal(NULL, TRUE, &lkbyt)))
            {
                ui_message_box(MB_ICONERROR, IDS_LOADFAIL1, hr);
                return 1;
            }

            try
            {
                ULARGE_INTEGER bufsz{ 0 };
                bufsz.QuadPart = buf.size();
                lkbyt->SetSize(bufsz);
                lkbyt->WriteAt(ULARGE_INTEGER{ 0 }, buf.data(), bufsz.LowPart, &bufsz.LowPart);
            }
            catch (const _com_error& exc)
            {
                ui_message_box(MB_ICONERROR, IDS_LOADFAIL, exc.Error(), exc.Description());
            }

            if (FAILED(hr = StgOpenStorageOnILockBytes(lkbyt, NULL, STGM_READ | STGM_SHARE_EXCLUSIVE, NULL, 0, &stg)))
            {
                ui_message_box(MB_ICONERROR, IDS_LOADFAIL1, hr);
                return 1;
            }
        }

        VBE::IVbaProjectPtr proj;
        try
        {
            proj = VbaHelper::GetInstance().LoadProject(stg);
        }
        catch (const _com_error& exc)
        {
            ui_message_box(MB_ICONERROR, IDS_LOADFAIL, exc.Error(), exc.Description());
        }

        // Show notification icon from here on
        NotifyIcon __notifyIcon(filename);

        // XXX: Better way to do this
        long exists, empty;
        proj->GetVbaProcs()->DoesProcExist(NULL, (LPWSTR)L"Main", &exists, &empty);
        try
        {
            if (exists)
            {
                proj->ExecuteLine((LPWSTR)L"Call Main(): While UserForms.Count > 0: DoEvents: Wend");
            }
            else
            {
                proj->ExecuteLine((LPWSTR)L"Call Main.Show(): While UserForms.Count > 0: DoEvents: Wend");
            }
        }
        catch (const _com_error& exc)
        {
            // This is likely a language-level error, so show the same dialog as in VBE
            VbaHelper::GetInstance().GetVba()->ShowError(exc.Error());
        }

        // XXX: ExecuteLine pushes its own message loop
        //ret = (int)VbaHelper::GetInstance().MessageLoop(); // HACK: See VbaCompManagerSite::raw_PushMessageLoop

        proj->Close();
    }
    catch (const _com_error& exc)
    {
        ui_message_box(MB_ICONERROR, IDS_COMERROR, exc.Error(), (wchar_t*)exc.Description());
    }

    return ret;
}

