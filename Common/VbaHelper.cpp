#include "stdafx.h"
#include "VbaHelper.h"

#include "ui.h"

#include "VbaSite.h"
#include "VbaProjectSite.h"
#include "VbaCompManagerSite.h"

VbaHelper VbaHelper::g_Instance;

VbaHelper::VbaHelper(const VbaHelperParams& params)
{
    LCID lcid = params.lcid;
    if (lcid == LOCALE_USER_DEFAULT)
    {
        lcid = GetUserDefaultLCID();
    }
    else if (lcid == LOCALE_SYSTEM_DEFAULT)
    {
        lcid = GetSystemDefaultLCID();
    }

    m_vbaSite = VBE::IVbaSitePtr(VbaSite::CreateInstance(params.headless, lcid), false);
    m_vbaCompManagerSite = VBE::IVbaCompManagerSitePtr(VbaCompManagerSite::CreateInstance(), false);

    VBE::VBAINITINFO vbaInitInfo{ sizeof(vbaInitInfo) };
    vbaInitInfo.dwFlags = VBE::VBAINITINFO_MessageLoopWide /* | VBE::VBAINITINFO_EnableDigitalSignatures | VBE::VBAINITINFO_SupportVbaDesigners */;
    if (params.headless)
    {
        vbaInitInfo.dwFlags |= VBE::VBAINITINFO_DisableIDE;
    }
    if (params.mdehack)
    {
        // NOTE: 0x4000 combined with VBACOOKIE_OFFICE sets VBE7!g_fHostIsAccess, allowing VBE UI to properly recognize MDE/ACCDE VBA projects
        // However this prevents IVbaProject::Load to load normal SA projects from disk; see VBE7!Project::ProcessProperties
        vbaInitInfo.dwFlags |= 0x4000;
    }

    vbaInitInfo.pVbaSite = m_vbaSite;
    vbaInitInfo.pVbaCompManagerSite = m_vbaCompManagerSite;

    vbaInitInfo.hwndHost = NULL; // We don't currently have a host window
    vbaInitInfo.hinstHost = GetModuleHandleW(NULL);
    vbaInitInfo.lpstrName = ui_load_string(params.nameid);
    //vbaInitInfo.lpstrRegKey = (LPWSTR)L"Common";
    vbaInitInfo.lcidUserInterface = lcid; // NOTE: Must be an actual LCID; 0 or pseudo LCIDs (e.g. LOCALE_SYSTEM_DEFAULT) won't work

    auto hr = VbaInitialize(params.cookie, &vbaInitInfo, &m_vba);
    if (FAILED(hr))
    {
        _com_raise_error(hr);
    }

    ((VbaCompManagerSite*)m_vbaCompManagerSite.GetInterfacePtr())->_SetVba(m_vba); // HACK: Find a better way to do this
}

VbaHelper::~VbaHelper()
{
    if (m_vba)
    {
        // NOTE: Detach here since VbaTerminate -> VBE7!DllVbeTerm -> VBE7!CVbe::Term calls CVbe::Release at end
        m_vba->PreTerminate();
        VbaTerminate(m_vba.Detach());
    }
}

VbaHelper& VbaHelper::operator=(VbaHelper&& rhs) noexcept
{
    if (this != &rhs)
    {
        m_vba = std::move(rhs.m_vba);
        m_vbaSite = std::move(rhs.m_vbaSite);
        m_vbaCompManagerSite = std::move(rhs.m_vbaCompManagerSite);
    }
    return *this;
}

void VbaHelper::Reset()
{
    try
    {
        m_vba->Reset();
    }
    catch (const _com_error& exc)
    {
        if (exc.Error() != VBE::VBA_E_RESET)
            throw;

        m_vba->GetCompManager()->InitiateReset();
    }
}

long VbaHelper::MessageLoop()
{
    HRESULT hr;
    long fAborted = 0;
    if (FAILED(hr = m_vbaCompManagerSite->PushMessageLoop((VBE::VBALOOPREASON)0, &fAborted)))
    {
        _com_raise_error(hr);
    }
    return fAborted;
}

VBE::IVbaProjectPtr VbaHelper::LoadProject(IStoragePtr pStg, bool initnew)
{
    VBE::IVbaProjectSitePtr vbaProjectSite(VbaProjectSite::CreateInstance(), false);
    auto vbaProject = m_vba->CreateProject(VBE::VBAPROJFLAG_Hidden, vbaProjectSite);
    try
    {
        if (initnew)
        {
            IPersistStoragePtr(vbaProject)->InitNew(pStg);
        }
        else
        {
            IPersistStoragePtr(vbaProject)->Load(pStg);
        }
    }
    catch (...)
    {
        vbaProject->Close();
        throw;
    }
    return vbaProject;
}

