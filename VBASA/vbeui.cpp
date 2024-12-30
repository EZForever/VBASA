#include "stdafx.h"
#include "resource.h"
#include "vbeui.h"

#include "ui.h"
#include "build.h"
#include "lickey.h"
#include "evalkey.h"
#include "xcomdef.h"

#include "VbaHelper.h"
#include "CommandBarControlEvents.h"

static void enumerate_command_bar(FILE* fp, Office::CommandBarPtr bar, int indent = 0)
{
    int bar_index = 0;
    try
    {
        bar_index = bar->Index;
    }
    catch (...)
    {
        // ...
    }

    for (int i = 0; i < indent; i++)
    {
        fputwc(L'\t', fp);
    }
    fwprintf(fp, L"%c#%d\t%d\t%s\t%s\n", L'B', bar_index, bar->Id, (wchar_t*)bar->Context, (wchar_t*)bar->Name);

    for (int i = 1; i < bar->Controls->Count; i++)
    {
        auto ctl = bar->Controls->Item[i];
        Office::CommandBarPopupPtr popup;
        try
        {
            popup = ctl;
        }
        catch (...)
        {
            // ...
        }

        for (int i = 0; i < indent + 1; i++)
        {
            fputwc(L'\t', fp);
        }
        fwprintf(fp, L"%c#%d\t%d\t%s\n", popup ? L'P' : L'C', ctl->Index, ctl->Id, (wchar_t*)ctl->Caption);
        if (popup)
        {
            enumerate_command_bar(fp, popup->CommandBar, indent + 2);
        }
    }
}

template<class IIID>
static DWORD advise_sink(IConnectionPointContainerPtr container, _com_ptr_t<IIID> sink)
{
    IConnectionPointPtr point;
    auto hr = container->FindConnectionPoint(IIID::GetIID(), &point);
    if (FAILED(hr))
    {
        _com_raise_error(hr);
    }

    DWORD cookie;
    hr = point->Advise(sink, &cookie);
    if (FAILED(hr))
    {
        _com_raise_error(hr);
    }

    return cookie;
}

static VBIDE::_VBProjectPtr get_current_project(VBIDE::VBEPtr vbe)
{
    auto proj = vbe->ActiveVBProject;
    if (!proj)
    {
        // This really should not happen since the button is greyed out if no project is open
        ui_message_box(MB_ICONWARNING, IDS_NOPROJ);
        return nullptr;
    }

    // XXX: This whole process is *tedious*, but since we have no access to the underlying IVbaProject this have to do

    // Look for Sub Main in code modules
    bool found = false;
    for (int i = proj->VBComponents->Count; i > 0; i--)
    {
        auto comp = proj->VBComponents->Item(i);
        if (comp->Type != VBIDE::vbext_ComponentType::vbext_ct_StdModule)
        {
            continue;
        }

        long line;
        try
        {
            line = comp->CodeModule->GetProcBodyLine(L"Main", VBIDE::vbext_ProcKind::vbext_pk_Proc);
        }
        catch (const _com_error&)
        {
            continue;
        }

        comp->CodeModule->CodePane->Show();
        comp->CodeModule->CodePane->SetSelection(line, 1, line, 1);
        found = true;
        break;
    }

    // If not found any, look for a UserForm named Main
    if (!found)
    {
        for (int i = proj->VBComponents->Count; i > 0; i--)
        {
            auto comp = proj->VBComponents->Item(i);
            if (comp->Type == VBIDE::vbext_ComponentType::vbext_ct_MSForm && comp->Name == _bstr_t(L"Main"))
            {
                comp->Activate();
                found = true;
                break;
            }
        }
    }

    // If still not found, warn the user
    if (!found)
    {
        ui_message_box(MB_ICONWARNING, IDS_NOMAIN);
        return nullptr;
    }

    return proj;
}

static void get_build_name(_bstr_t& buildname, bool& strip)
{
    HRESULT hr;

    // TODO: Save filter selected filter idx
    // TODO: Other build types
    auto bstrStdExe = ui_load_string(IDS_BLDSTDEXE);
    //auto bstrAny = ui_load_string(IDS_BLDANY); // XXX: This will cause confusion in later steps
    COMDLG_FILTERSPEC rgFilterSpecs[] = {
        { bstrStdExe.GetBSTR(), L"*.exe" },
        //{ bstrAny.GetBSTR(), L"*.*" }
    };

    IFileSaveDialogPtr dlg(CLSID_FileSaveDialog);
    _XCOM_RAISE_IF_FAILED(dlg->SetTitle(ui_load_string(IDS_BLDTITLE)));
    _XCOM_RAISE_IF_FAILED(dlg->SetFileTypes(_countof(rgFilterSpecs), rgFilterSpecs));
    _XCOM_RAISE_IF_FAILED(dlg->SetFileName(buildname));

    IFileDialogCustomizePtr dlgCustomize(dlg);
    dlgCustomize->AddCheckButton(1, ui_load_string(IDS_BLDSRP), strip);
    // TODO: Register on Build for ActiveX types?

    if (FAILED(hr = dlg->Show(g_hwndMainWindow)))
    {
        // User cancelled the dialog. Abort
        buildname.Assign(NULL);
        return;
    }

    IShellItemPtr item;
    _XCOM_RAISE_IF_FAILED(dlg->GetResult(&item));

    LPWSTR pszName;
    _XCOM_RAISE_IF_FAILED(item->GetDisplayName(SIGDN_FILESYSPATH, &pszName));
    buildname = pszName;
    CoTaskMemFree(pszName);

    BOOL fStripChecked;
    _XCOM_RAISE_IF_FAILED(dlgCustomize->GetCheckButtonState(1, &fStripChecked));
    strip = fStripChecked;
    if (strip && ui_message_box(MB_ICONWARNING | MB_YESNO, IDS_BLDSRPWARN) != IDYES)
    {
        // User regreted on building SRP executable
        buildname.Assign(NULL);
        return;
    }
}

bool init_vbeui(VBIDE::VBEPtr vbe)
{
#if 0
    FILE* fp;
    _wfopen_s(&fp, L"command_bars.log", L"w,ccs=UTF-8");
    for (int i = 1; i <= vbe->CommandBars->Count; i++)
    {
        enumerate_command_bar(fp, vbe->CommandBars->Item[i]);
    }
    fclose(fp);
#endif

    // XXX: VBE6EXT has HWND stored in LONG; let's hope the value never goes beyond 32-bit limit
    g_hwndMainWindow = (HWND)(UINT_PTR)vbe->GetMainWindow()->HWnd;

    VBIDE::_dispCommandBarControlEventsPtr events(CommandBarControlEvents::CreateInstance(), false);

    // NOTE: Removed entries can still be found in toolbar customization dialog
    // NOTE: Property modifications are permanent

    // NOTE: OnAction strings are ignored when using VBEUI; see VBEUI!TBC::FRunOnAction -> VBE7!CMsoUser::FRunMacro

    // Repurpose broken "Build Project.DLL..." menu entry
    // MenuBar#32768->File#30002->Popup#32769->BuildProject#746[-1]
    {
        Office::CommandBarPopupPtr file_popup = vbe->CommandBars->FindControl(Office::MsoControlType::msoControlPopup, 30002);
        for (int i = file_popup->CommandBar->Controls->Count; i > 0; i--)
        {
            auto ctl = file_popup->CommandBar->Controls->Item[i];
            if (ctl->Id == 746)
            {
                //ctl->Delete(VARIANT_TRUE);
                advise_sink(vbe->Events->CommandBarEvents[ctl], events); // XXX: Never unadvised
                break;
            }
        }
    }

    // Repurpose broken "Run Project" toolbar icons
    // MenuBar#32768->Run#30012->Popup#32775->RunProject#5415
    // Standard#32815->RunProject#5415
    // Debug#32817->RunProject#5415
    {
        auto ctls = vbe->CommandBars->FindControls(Office::MsoControlType::msoControlButton, 5415);
        for (int i = ctls->Count; i > 0; i--)
        {
            auto ctl = ctls->Item[i];
            advise_sink(vbe->Events->CommandBarEvents[ctl], events); // XXX: Never unadvised
        }
    }

    // Remove "Design Mode" toolbar icons (will leave no holes since "Run Project" appears just after it in every case)
    // Standard#32815->DesignMode#212
    // Debug#32817->DesignMode#212
    // MenuBar#32768->Run#30012->Popup#32775->DesignMode#212
    {
        auto ctls = vbe->CommandBars->FindControls(Office::MsoControlType::msoControlButton, 212);
        for (int i = ctls->Count; i > 0; i--)
        {
            ctls->Item[i]->Delete(VARIANT_TRUE);
        }
    }

    // Remove "Application View" toolbar icons
    // Standard#32815->Application#106 (will leave a hole)
    // MenuBar#32768->View#30004->Popup#32771->Application#106
    {
        auto ctls = vbe->CommandBars->FindControls(Office::MsoControlType::msoControlButton, 106);
        for (int i = ctls->Count; i > 0; i--)
        {
            ctls->Item[i]->Delete(VARIANT_TRUE);
        }
    }

    // Fill the hole left by removing the application view, with a custom "New Project" button
    {
        Office::CommandBarPtr standard_bar;
        for (int i = vbe->CommandBars->Count; i > 0; i--)
        {
            auto bar = vbe->CommandBars->Item[i];
            if (bar->Id == 32815)
            {
                standard_bar = bar;
                break;
            }
        }

        Office::CommandBarPopupPtr file_popup = vbe->CommandBars->FindControl(Office::MsoControlType::msoControlPopup, 30002);
        Office::_CommandBarButtonPtr new_proj_ctl;
        for (int i = file_popup->CommandBar->Controls->Count; i > 0; i--)
        {
            auto ctl = file_popup->CommandBar->Controls->Item[i];
            if (ctl->Id == 746 && ctl->Index == 1)
            {
                new_proj_ctl = ctl;
                break;
            }
        }

        if (standard_bar && new_proj_ctl)
        {
            Office::_CommandBarButtonPtr ctl = standard_bar->Controls->Add(Office::MsoControlType::msoControlButton, vtMissing, vtMissing, 1, VARIANT_TRUE);
            ctl->Caption = new_proj_ctl->Caption;
            ctl->TooltipText = new_proj_ctl->Caption;
            //ctl->FaceId = 2557; // ProjectResourceManager
            ctl->FaceId = 1695; // VisualBasic

            ctl->OnAction = L"!<VBASA>";
            ctl->Parameter = L"NewProject";
            advise_sink(vbe->Events->CommandBarEvents[ctl], events); // XXX: Never unadvised
        }
    }

    if(wcsstr(GetCommandLineW(), L"Doing_things_wrong"))
    {
        wchar_t szCaption[64];
        auto cbCaption = LoadStringW(GetModuleHandleW(L"VBE7INTL.DLL"), 18089, szCaption, _countof(szCaption));
        Office::CommandBarPopupPtr help_popup = vbe->CommandBars->FindControl(Office::MsoControlType::msoControlPopup, 30010);
        if (cbCaption && help_popup)
        {
            Office::_CommandBarButtonPtr ctl = help_popup->CommandBar->Controls->Add(Office::MsoControlType::msoControlButton, vtMissing, vtMissing, 2, VARIANT_TRUE);
            ctl->Caption = szCaption;
            ctl->TooltipText = szCaption;
            //ctl->ShortcutText = L"Shift+F1";
            ctl->ShortcutText = L"~";
            ctl->FaceId = 59; // SmileFace
            //ctl->FaceId = 487; // Information

            ctl->OnAction = L"!<VBASA>";
            ctl->Parameter = L"By~f1GS1Rctuxeb";
            advise_sink(vbe->Events->CommandBarEvents[ctl], events); // XXX: Never unadvised
        }
    }

    return true;
}

void vbe_new_project(VBIDE::VBEPtr vbe)
{
    Office::CommandBarPopupPtr file_popup = vbe->CommandBars->FindControl(Office::MsoControlType::msoControlPopup, 30002);
    Office::_CommandBarButtonPtr new_proj_ctl;
    for (int i = file_popup->CommandBar->Controls->Count; i > 0; i--)
    {
        auto ctl = file_popup->CommandBar->Controls->Item[i];
        if (ctl->Id == 746 && ctl->Index == 1)
        {
            new_proj_ctl = ctl;
            break;
        }
    }

    new_proj_ctl->Execute();
}

void vbe_build_project(VBIDE::VBEPtr vbe)
{
//#ifndef _DEBUG
    if (!_wcsicmp(g_bstrLicKey, g_wszVbaEvalLicKey))
    {
        ui_message_box(MB_ICONWARNING, IDS_BLDLIC);
        return;
    }
//#endif

    auto proj = get_current_project(vbe);
    if (!proj)
    {
        return;
    }

    // Check if the project is current on disk
    // XXX: This is not necessary in theory since proj->SaveAs could force that; however that sliently changed the file opened
    _bstr_t filename;
    try
    {
        filename = proj->FileName;
    }
    catch (const _com_error&)
    {
        // The project has never been saved; help the user by opening the save dialog
        ui_message_box(MB_ICONWARNING, IDS_BLDSAVE);

        Office::_CommandBarButtonPtr ctl = vbe->CommandBars->FindControl(Office::MsoControlType::msoControlButton, 3);
        if (ctl)
        {
            ctl->Execute();
        }
        
        return;
    }

    if (!proj->Saved)
    {
        // The project was once saved but is now dirty; just ping the user, do not auto save
        ui_message_box(MB_ICONWARNING, IDS_BLDSAVE);
        return;
    }

    _bstr_t buildname = proj->BuildFileName;
    bool strip = false;

    // Default build name always ends with .DLL
    // We don't support building DLLs for now so silently change that
    // If the user just chooses to build a DLL then the extension name would be in lower case so no problem there
    // NOTE: BSTRs don't end with L'\0'
    auto extname = buildname.GetBSTR() + buildname.length() - 4;
    if (!memcmp(extname, L".DLL", 4 * sizeof(WCHAR)))
    {
        memcpy(extname, L".exe", 4 * sizeof(WCHAR));
    }

    // Show the Save As dialog, and abort if user cancelled the build
    get_build_name(buildname, strip);
    if (buildname.length() == 0)
    {
        return;
    }

    build_project(filename, buildname, strip);

    // Save the modified build name to the project
    if (buildname != proj->BuildFileName)
    {
        proj->BuildFileName = buildname;
    }
}

void vbe_run_project(VBIDE::VBEPtr vbe)
{
    auto proj = get_current_project(vbe);
    if (!proj)
    {
        return;
    }

    // Do not try to run the same project twice
    if (proj->Mode != VBIDE::vbext_VBAMode::vbext_vm_Design)
    {
        ui_message_box(MB_ICONWARNING, IDS_RUNALDY);
        return;
    }

    // Cue the Run Macro button
    Office::_CommandBarButtonPtr ctl = vbe->CommandBars->FindControl(Office::MsoControlType::msoControlButton, 186);
    if (ctl)
    {
        ctl->Execute();
    }
}

