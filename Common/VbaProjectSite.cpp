#include "stdafx.h"
#include "VbaProjectSite.h"

VBE::IVbaProjectSite* VbaProjectSite::CreateInstance()
{
    auto ptr = new VbaProjectSite();
    ptr->AddRef();
    return ptr;
}

// --- Implements IVbaProjectSite->IUnknown

STDMETHODIMP VbaProjectSite::QueryInterface(REFIID riid, void** ppvObject)
{
    if (!ppvObject)
        return E_POINTER;

    if (riid != IID_IUnknown && riid != __uuidof(VBE::IVbaProjectSite))
    {
        *ppvObject = nullptr;
        return E_NOINTERFACE;
    }

    this->AddRef();
    *ppvObject = this;
    return S_OK;
}

STDMETHODIMP_(ULONG) VbaProjectSite::AddRef()
{
    return _InterlockedIncrement(&m_refcnt);
}

STDMETHODIMP_(ULONG) VbaProjectSite::Release()
{
    LONG lref = _InterlockedDecrement(&m_refcnt);
    if (lref == 0)
        delete this;
    return lref;
}

// --- Implements IVbaProjectSite

// NOTE: This is a stripped-down replica of VBE7!CProjSite, which standard SA project uses
// Removed loading/saving ability

STDMETHODIMP VbaProjectSite::raw_GetOwnerWindow(wireHWND* phwnd) {
    *phwnd = NULL; // XXX: VBE7!CProjSite::GetOwnerWindow returns VBE7!Main_hwndMainMenu
    return S_OK;
}

STDMETHODIMP VbaProjectSite::raw_Notify(VBE::VBAPROJSITENOTIFY vbapsn) {
    // XXX: VBE7!CProjSite::Notify(VBAPSN_Activate) calls VbaSite::Notify(VBAHN_Activate)
	return S_OK;
}

STDMETHODIMP VbaProjectSite::raw_GetObjectOfName(LPWSTR lpstrName, VBE::IDocumentSite** ppDocSite) {
    *ppDocSite = NULL;
	return E_NOTIMPL;
}

STDMETHODIMP VbaProjectSite::raw_LockProjectOwner(long fLock, long fLastUnlockCloses) {
	return S_OK;
}

STDMETHODIMP VbaProjectSite::raw_RequestSave(wireHWND hwndOwner) {
	return S_OK;
}

STDMETHODIMP VbaProjectSite::raw_ReleaseModule(VBE::IVbaProjItem* pVbaProjItem) {
	return S_OK;
}

STDMETHODIMP VbaProjectSite::raw_ModuleChanged(VBE::IVbaProjItem* pVbaProjItem) {
	return S_OK;
}

STDMETHODIMP VbaProjectSite::raw_DeletingProjItem(VBE::IVbaProjItem* pVbaProjItem) {
	return S_OK;
}

STDMETHODIMP VbaProjectSite::raw_ModeChange(VBE::VBA_PROJECT_MODE pvbaprojmode) {
	return S_OK;
}

STDMETHODIMP VbaProjectSite::raw_ProcChanged(VBE::VBA_PROC_CHANGE_INFO* pvbapcinfo) {
	return S_OK;
}

STDMETHODIMP VbaProjectSite::raw_ReleaseInstances(VBE::IVbaProjItem* pVbaProjItem) {
	return S_OK;
}

STDMETHODIMP VbaProjectSite::raw_ProjItemCreated(VBE::IVbaProjItem* pVbaProjItem) {
	return S_OK;
}

STDMETHODIMP VbaProjectSite::raw_GetMiniBitmapGuid(VBE::IVbaProjItem* pVbaProjItem, GUID* pguidMiniBitmap) {
	return E_NOTIMPL;
}

STDMETHODIMP VbaProjectSite::raw_GetSignature(void** ppvSignature, unsigned long* pcbSign) {
	return E_NOTIMPL;
}

STDMETHODIMP VbaProjectSite::raw_PutSignature(void* pvSignature, unsigned long cbSign) {
	return E_NOTIMPL;
}

STDMETHODIMP VbaProjectSite::raw_NameChanged(VBE::IVbaProjItem* pVbaProjItem, LPWSTR lpstrOldName) {
	return S_OK;
}

STDMETHODIMP VbaProjectSite::raw_ModuleDirtyChanged(VBE::IVbaProjItem* pVbaProjItem, long fDirty) {
	return S_OK;
}

STDMETHODIMP VbaProjectSite::raw_CreateDocClassInstance(VBE::IVbaProjItem* pVbaProjItem, GUID clsidDocClass, IUnknown* punkOuter, long fIsPredeclared, IUnknown** ppunkInstance) {
	return E_NOTIMPL;
}

