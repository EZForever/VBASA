#pragma once
#include "vbe.h"

class VbaProjectSite final : public VBE::IVbaProjectSite
{
private:
    ULONG m_refcnt = 0;

    VbaProjectSite() = default;
    VbaProjectSite(VbaProjectSite&&) = delete;

    ~VbaProjectSite() = default;

public:
    static VBE::IVbaProjectSite* CreateInstance();

    // --- Implements IVbaProjectSite->IUnknown

    STDMETHODIMP QueryInterface(REFIID riid, void** ppvObject);
    STDMETHODIMP_(ULONG) AddRef();
    STDMETHODIMP_(ULONG) Release();

    // --- Implements IVbaProjectSite

    STDMETHODIMP raw_GetOwnerWindow(wireHWND* phwnd);
    STDMETHODIMP raw_Notify(VBE::VBAPROJSITENOTIFY vbapsn);
    STDMETHODIMP raw_GetObjectOfName(LPWSTR lpstrName, VBE::IDocumentSite** ppDocSite);
    STDMETHODIMP raw_LockProjectOwner(long fLock, long fLastUnlockCloses);
    STDMETHODIMP raw_RequestSave(wireHWND hwndOwner);
    STDMETHODIMP raw_ReleaseModule(VBE::IVbaProjItem* pVbaProjItem);
    STDMETHODIMP raw_ModuleChanged(VBE::IVbaProjItem* pVbaProjItem);
    STDMETHODIMP raw_DeletingProjItem(VBE::IVbaProjItem* pVbaProjItem);
    STDMETHODIMP raw_ModeChange(VBE::VBA_PROJECT_MODE pvbaprojmode);
    STDMETHODIMP raw_ProcChanged(VBE::VBA_PROC_CHANGE_INFO* pvbapcinfo);
    STDMETHODIMP raw_ReleaseInstances(VBE::IVbaProjItem* pVbaProjItem);
    STDMETHODIMP raw_ProjItemCreated(VBE::IVbaProjItem* pVbaProjItem);
    STDMETHODIMP raw_GetMiniBitmapGuid(VBE::IVbaProjItem* pVbaProjItem, GUID* pguidMiniBitmap);
    STDMETHODIMP raw_GetSignature(void** ppvSignature, unsigned long* pcbSign);
    STDMETHODIMP raw_PutSignature(void* pvSignature, unsigned long cbSign);
    STDMETHODIMP raw_NameChanged(VBE::IVbaProjItem* pVbaProjItem, LPWSTR lpstrOldName);
    STDMETHODIMP raw_ModuleDirtyChanged(VBE::IVbaProjItem* pVbaProjItem, long fDirty);
    STDMETHODIMP raw_CreateDocClassInstance(VBE::IVbaProjItem* pVbaProjItem, GUID clsidDocClass, IUnknown* punkOuter, long fIsPredeclared, IUnknown** ppunkInstance);
};

