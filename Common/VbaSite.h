#pragma once
#include "vbe.h"

class VbaSite final : public VBE::IVbaSite
{
private:
    ULONG m_refcnt = 0;
    bool m_headless;
    LCID m_lcid;

    VbaSite(bool headless, LCID lcid)
        : m_headless(headless), m_lcid(lcid) {}
    VbaSite(VbaSite&&) = delete;

    ~VbaSite() = default;

public:
    static VBE::IVbaSite* CreateInstance(bool headless, LCID lcid);

    // --- Implements VbaSite->IUnknown

    STDMETHODIMP QueryInterface(REFIID riid, void** ppvObject);
    STDMETHODIMP_(ULONG) AddRef();
    STDMETHODIMP_(ULONG) Release();

    // --- Implements VbaSite

    STDMETHODIMP raw_GetOwnerWindow(wireHWND* phwnd);
    STDMETHODIMP raw_Notify(VBE::VBASITENOTIFY vbahnotify);
    STDMETHODIMP raw_FindFile(LPWSTR lpstrFileName, BSTR* pbstrFullPath);
    STDMETHODIMP raw_GetMiniBitmap(GUID* rgguid, wireHBITMAP* phbmp, unsigned long* pcrMask);
    STDMETHODIMP raw_CanEnterBreakMode(VBE::VBA_BRK_INFO* pbrkinfo);
    STDMETHODIMP raw_SetStepMode(VBE::VBASTEPMODE vbastepmode);
    STDMETHODIMP raw_ShowHide(long fVisible);
    STDMETHODIMP raw_InstanceCreated(IUnknown* punkInstance);
    STDMETHODIMP raw_HostCheckReference(long fSave, GUID* rgguid, UINT* puVerMaj, UINT* puVerMin);
    STDMETHODIMP raw_GetIEVersion(LONG* plMajVer, LONG* plMinVer, long* pfCanInstall);
    STDMETHODIMP raw_UseIEFeature(LONG lMajVer, LONG lMinVer, wireHWND hwndOwner);
    STDMETHODIMP raw_ShowHelp(wireHWND hwnd, LPWSTR szHelp, UINT wCmd, DWORD dwHelpContext, long fWinHelp);
    STDMETHODIMP raw_OpenProject(wireHWND hwndOwner, VBE::IVbaProject* pVbaRefProj, LPWSTR lpstrFileName, VBE::IVbaProject** ppVbaProj);
};

