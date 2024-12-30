#include "stdafx.h"
#include "VbaCompManagerSite.h"

VBE::IVbaCompManagerSite* VbaCompManagerSite::CreateInstance()
{
    auto ptr = new VbaCompManagerSite();
    ptr->AddRef();
    return ptr;
}

void VbaCompManagerSite::_SetVba(VBE::IVbaPtr vba)
{
    m_vba = vba;
}

// --- Implements IVbaCompManagerSite->IUnknown

STDMETHODIMP VbaCompManagerSite::QueryInterface(REFIID riid, void** ppvObject)
{
    if (!ppvObject)
        return E_POINTER;

    if (riid != IID_IUnknown && riid != __uuidof(VBE::IVbaCompManagerSite))
    {
        *ppvObject = nullptr;
        return E_NOINTERFACE;
    }

    this->AddRef();
    *ppvObject = this;
    return S_OK;
}

STDMETHODIMP_(ULONG) VbaCompManagerSite::AddRef()
{
    return _InterlockedIncrement(&m_refcnt);
}

STDMETHODIMP_(ULONG) VbaCompManagerSite::Release()
{
    LONG lref = _InterlockedDecrement(&m_refcnt);
    if (lref == 0)
        delete this;
    return lref;
}

// --- Implements IVbaCompManagerSite

STDMETHODIMP VbaCompManagerSite::raw_Notify(VBE::VBACMSITENOTIFY vbacmsn) {
    return S_OK;
}

STDMETHODIMP VbaCompManagerSite::raw_EnterState(VBE::VBACOMPSTATE vbacompstate, long fEnter) {
    return S_OK;
}

// Adapted from VBA6SDK ApcCpp.h, ApcHost_OnMessageLoop()
STDMETHODIMP VbaCompManagerSite::raw_PushMessageLoop(VBE::VBALOOPREASON vbaloopreason, long* pfAborted) {
    // NOTE: msgloop == 0 -> First message loop on this thread, should handle QueryTerminate and return value
    // vbaloopreason == 0 -> Message loop pushed by host (see VbaHelper::MessageLoop), for use with ContinueMessageLoop
    auto msgloop = _InterlockedIncrement(&m_msgloop);
    long ret = 0;
    try
    {
        auto vbaCompManager = m_vba->GetCompManager();
        while (vbaCompManager->ContinueMessageLoop(vbaloopreason == 0))
        {
            MSG msg;
            if (!PeekMessageW(&msg, NULL, WM_NULL, WM_NULL, PM_REMOVE))
            {
                bool idling = true;
                while (idling)
                {
                    for (auto flag : { VBE::VBAIDLEFLAG_Periodic, VBE::VBAIDLEFLAG_Priority, VBE::VBAIDLEFLAG_NonPeriodic })
                    {
                        if (vbaCompManager->DoIdle(flag))
                        {
                            break;
                        }
                        idling = false;
                    }
                }

                vbaCompManager->OnWaitForMessage();
                if (!PeekMessageW(&msg, NULL, WM_NULL, WM_NULL, PM_NOREMOVE))
                {
                    WaitMessage();
                }

                continue;
            }

            if (msg.message == WM_QUIT)
            {
                if (msgloop == 0)
                {
                    if (m_vba->QueryTerminate())
                    {
                        ret = (long)msg.wParam;
                        break;
                    }
                    else
                    {
                        continue;
                    }
                }
                else
                {
                    PostQuitMessage((int)msg.wParam);
                    break;
                }
            }

            if (!vbaCompManager->PreTranslateMessage(&msg))
            {
                TranslateMessage(&msg);
                DispatchMessageW(&msg);
            }
        }
    }
    catch (const _com_error& exc)
    {
        _InterlockedDecrement(&m_msgloop);

        *pfAborted = TRUE;
        return exc.Error();
    }

    if (msgloop == 0)
        *pfAborted = ret; // HACK: Pass return value to host
    else
        *pfAborted = FALSE;
    return S_OK;
}

STDMETHODIMP VbaCompManagerSite::raw_ContinueIdle(long* pfContinue) {
    MSG msg;
    *pfContinue = !PeekMessageW(&msg, NULL, WM_NULL, WM_NULL, PM_NOREMOVE);
    return S_OK;
}

STDMETHODIMP VbaCompManagerSite::raw_DoIdle() {
    return S_OK;
}

