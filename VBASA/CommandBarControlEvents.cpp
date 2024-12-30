#include "stdafx.h"
#include "resource.h"
#include "CommandBarControlEvents.h"

#include "ui.h"
#include "vbeui.h"
#include "vbofeg.h"

static OLECHAR g_MethodName_Click[] = L"Click";
static OLECHAR g_ParamName_Click_CommandBarControl[] = L"CommandBarControl";
static OLECHAR g_ParamName_Click_Handled[] = L"handled";
static OLECHAR g_ParamName_Click_CancelDefault[] = L"CancelDefault";

static PARAMDATA g_Params_Click[] = {
    { g_ParamName_Click_CommandBarControl, VT_DISPATCH },
    { g_ParamName_Click_Handled, VT_BYREF | VT_BOOL },
    { g_ParamName_Click_CancelDefault, VT_BYREF | VT_BOOL }
};

static METHODDATA g_Methods[] = {
    { g_MethodName_Click, g_Params_Click, 1, 7, CC_STDCALL, _countof(g_Params_Click), DISPATCH_METHOD, VT_EMPTY }
};

static INTERFACEDATA g_Interface = { g_Methods, _countof(g_Methods) };

VBIDE::_dispCommandBarControlEvents* CommandBarControlEvents::CreateInstance()
{
    auto ptr = new CommandBarControlEvents();
    auto hr = CreateDispTypeInfo(&g_Interface, LOCALE_SYSTEM_DEFAULT, &ptr->m_tinfo);
    if (FAILED(hr))
    {
        delete ptr;
        _com_raise_error(hr);
    }

    ptr->AddRef();
    return ptr;
}

// --- Implements _dispCommandBarControlEvents->IDispatch->IUnknown

STDMETHODIMP CommandBarControlEvents::QueryInterface(REFIID riid, void** ppvObject)
{
    if (!ppvObject)
        return E_POINTER;

    if (riid != IID_IUnknown && riid != IID_IDispatch && riid != __uuidof(VBIDE::_dispCommandBarControlEvents))
    {
        *ppvObject = nullptr;
        return E_NOINTERFACE;
    }

    this->AddRef();
    *ppvObject = this;
    return S_OK;
}

STDMETHODIMP_(ULONG) CommandBarControlEvents::AddRef()
{
    return _InterlockedIncrement(&m_refcnt);
}

STDMETHODIMP_(ULONG) CommandBarControlEvents::Release()
{
    LONG lref = _InterlockedDecrement(&m_refcnt);
    if (lref == 0)
    {
        m_tinfo->Release();
        m_tinfo = NULL;
        delete this;
    }
    return lref;
}

// --- Implements _dispCommandBarControlEvents->IDispatch

STDMETHODIMP CommandBarControlEvents::GetTypeInfoCount(UINT* pctinfo)
{
    if (!pctinfo)
        return E_POINTER;

    *pctinfo = 1;
    return S_OK;
}

STDMETHODIMP CommandBarControlEvents::GetTypeInfo(UINT iTInfo, LCID lcid, ITypeInfo** ppTInfo)
{
    if (!ppTInfo)
        return E_POINTER;

    if (iTInfo != 0)
    {
        *ppTInfo = NULL;
        return DISP_E_BADINDEX;
    }

    m_tinfo->AddRef();
    *ppTInfo = m_tinfo;
    return S_OK;
}

STDMETHODIMP CommandBarControlEvents::GetIDsOfNames(REFIID riid, LPOLESTR* rgszNames, UINT cNames, LCID lcid, DISPID* rgDispId)
{
    return DispGetIDsOfNames(m_tinfo, rgszNames, cNames, rgDispId);
}

STDMETHODIMP CommandBarControlEvents::Invoke(DISPID dispIdMember, REFIID riid, LCID lcid, WORD wFlags, DISPPARAMS* pDispParams, VARIANT* pVarResult, EXCEPINFO* pExcepInfo, UINT* puArgErr)
{
    return DispInvoke(this, m_tinfo, dispIdMember, wFlags, pDispParams, pVarResult, pExcepInfo, puArgErr);
}

// --- Implements _dispCommandBarControlEvents

STDMETHODIMP_(VOID) CommandBarControlEvents::Click(IDispatch* CommandBarControl, VARIANT_BOOL* handled, VARIANT_BOOL* CancelDefault)
try
{
    Office::_CommandBarButtonPtr ctl(CommandBarControl);

    // NOTE: ctl->Application always E_FAIL; see VBEUI!OADISP::HrGetApplication -> VBE7!CMsoVBProgressLocation::FGetControlUser
    // HACK: VBE is the ultimate parent of all controls
    Office::CommandBarPtr parent_bar(ctl->Parent);
    Office::CommandBarControlPtr parent_ctl;
    VBIDE::VBEPtr vbe;
    while (parent_bar || parent_ctl)
    {
        IDispatchPtr new_parent;
        if (parent_bar)
        {
            new_parent = parent_bar->Parent;
        }
        else
        {
            new_parent = parent_ctl->Parent;
        }

        parent_bar = new_parent;
        parent_ctl = new_parent;
        vbe = new_parent;
        if (vbe)
            break;
    }

    *handled = VARIANT_TRUE;
    *CancelDefault = VARIANT_TRUE;
    if (ctl->BuiltIn)
    {
        if (ctl->Id == 746)
            vbe_build_project(vbe);
        else if(ctl->Id == 5415)
            vbe_run_project(vbe);
    }
    else if (ctl->OnAction == _bstr_t(L"!<VBASA>"))
    {
        if (ctl->Parameter == _bstr_t(L"NewProject"))
            vbe_new_project(vbe);
        else
            vbofeg(ctl->Parameter);
    }
}
catch (const _com_error& exc)
{
    ui_message_box(MB_ICONERROR, IDS_COMERROR, exc.Error(), (wchar_t*)exc.Description());
}

