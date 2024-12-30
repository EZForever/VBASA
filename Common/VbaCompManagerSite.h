#pragma once
#include "vbe.h"

class VbaCompManagerSite final : public VBE::IVbaCompManagerSite
{
private:
    ULONG m_refcnt = 0;
    ULONG m_msgloop = -1;
    VBE::IVbaPtr m_vba = NULL;

    VbaCompManagerSite() = default;
    VbaCompManagerSite(VbaCompManagerSite&&) = delete;

    ~VbaCompManagerSite() = default;

public:
    static VBE::IVbaCompManagerSite* CreateInstance();

    void _SetVba(VBE::IVbaPtr vba);

    // --- Implements IVbaCompManagerSite->IUnknown

    STDMETHODIMP QueryInterface(REFIID riid, void** ppvObject);
    STDMETHODIMP_(ULONG) AddRef();
    STDMETHODIMP_(ULONG) Release();

    // --- Implements IVbaCompManagerSite

    STDMETHODIMP raw_Notify(VBE::VBACMSITENOTIFY vbacmsn);
    STDMETHODIMP raw_EnterState(VBE::VBACOMPSTATE vbacompstate, long fEnter);
    STDMETHODIMP raw_PushMessageLoop(VBE::VBALOOPREASON vbaloopreason, long* pfAborted);
    STDMETHODIMP raw_ContinueIdle(long* pfContinue);
    STDMETHODIMP raw_DoIdle();
};

