#include "stdafx.h"
#include "FauxTypeLib2.h"

#include <initguid.h>

// {FB1E1EC0-A735-11D1-96E8-00A0C91110EB}
DEFINE_GUID(DIID__DPropPublicClassModule, 0xFB1E1EC0, 0xA735, 0x11D1, 0x96, 0xE8, 0x00, 0xA0, 0xC9, 0x11, 0x10, 0xEB);
// {25411250-A738-11D1-96E8-00A0C91110EB}
DEFINE_GUID(DIID__DPropPrivateClassModule, 0x25411250, 0xA738, 0x11D1, 0x96, 0xE8, 0x00, 0xA0, 0xC9, 0x11, 0x10, 0xEB);

ITypeLib2* FauxTypeLib2::CreateInstance(ITypeLib* tlib)
{
    if (!tlib)
        _com_raise_error(E_POINTER);

    HRESULT hr;
    ITypeLib2* tlib2 = NULL;
    if (FAILED(hr = tlib->QueryInterface(&tlib2)))
    {
        if (hr != E_NOINTERFACE)
            _com_raise_error(hr);
    }

    auto ptr = new FauxTypeLib2(tlib, tlib2);
    tlib->AddRef();
    ptr->AddRef();
    return ptr;
}

// --- Implements ITypeLib2->ITypeLib->IUnknown

STDMETHODIMP FauxTypeLib2::QueryInterface(REFIID riid, void** ppvObject)
{
    if (!ppvObject)
        return E_POINTER;

    if (riid != IID_IUnknown && riid != __uuidof(ITypeLib) && !(m_tlib2 && riid == __uuidof(ITypeLib2)))
    {
        *ppvObject = nullptr;
        return E_NOINTERFACE;
    }

    this->AddRef();
    *ppvObject = this;
    return S_OK;
}

STDMETHODIMP_(ULONG) FauxTypeLib2::AddRef()
{
    return _InterlockedIncrement(&m_refcnt);
}

STDMETHODIMP_(ULONG) FauxTypeLib2::Release()
{
    LONG lref = _InterlockedDecrement(&m_refcnt);
    if (lref == 0)
    {
        m_tlib->Release();
        if (m_tlib2)
            m_tlib2->Release();
        delete this;
    }
    return lref;
}

// --- Implements ITypeLib2->ITypeLib

STDMETHODIMP_(UINT) FauxTypeLib2::GetTypeInfoCount()
{
    return m_tlib->GetTypeInfoCount();
}

STDMETHODIMP FauxTypeLib2::GetTypeInfo(UINT index, ITypeInfo** ppTInfo)
{
    return m_tlib->GetTypeInfo(index, ppTInfo);
}

STDMETHODIMP FauxTypeLib2::GetTypeInfoType(UINT index, TYPEKIND* pTKind)
{
    return m_tlib->GetTypeInfoType(index, pTKind);
}

STDMETHODIMP FauxTypeLib2::GetTypeInfoOfGuid(REFGUID guid, ITypeInfo** ppTinfo)
{
    // VBE7!ErrGetPropertyDispatch only initializes CMProperties with type #5 (_DPropNameOnly) or #10 (_DPropPrivateClassModule)
    // VBE7!TLCGetTypeInfoNoRef then do a lookup of this internal ID to DIID in VBInternal typelib
    auto& new_guid = (guid == DIID__DPropPrivateClassModule) ? DIID__DPropPublicClassModule : guid;
    return m_tlib->GetTypeInfoOfGuid(new_guid, ppTinfo);
}

STDMETHODIMP FauxTypeLib2::GetLibAttr(TLIBATTR** ppTLibAttr)
{
    return m_tlib->GetLibAttr(ppTLibAttr);
}

STDMETHODIMP FauxTypeLib2::GetTypeComp(ITypeComp** ppTComp)
{
    return m_tlib->GetTypeComp(ppTComp);
}

STDMETHODIMP FauxTypeLib2::GetDocumentation(INT index, BSTR* pBstrName, BSTR* pBstrDocString, DWORD* pdwHelpContext, BSTR* pBstrHelpFile)
{
    return m_tlib->GetDocumentation(index, pBstrName, pBstrDocString, pdwHelpContext, pBstrHelpFile);
}

STDMETHODIMP FauxTypeLib2::IsName(LPOLESTR szNameBuf, ULONG lHashVal, BOOL* pfName)
{
    return m_tlib->IsName(szNameBuf, lHashVal, pfName);
}

STDMETHODIMP FauxTypeLib2::FindName(LPOLESTR szNameBuf, ULONG lHashVal, ITypeInfo** ppTInfo, MEMBERID* rgMemId, USHORT* pcFound)
{
    return m_tlib->FindName(szNameBuf, lHashVal, ppTInfo, rgMemId, pcFound);
}

STDMETHODIMP_(void) FauxTypeLib2::ReleaseTLibAttr(TLIBATTR* pTLibAttr)
{
    return m_tlib->ReleaseTLibAttr(pTLibAttr);
}

// --- Implements ITypeLib2

STDMETHODIMP FauxTypeLib2::GetCustData(REFGUID guid, VARIANT* pVarVal)
{
    if (!m_tlib2)
        return E_NOTIMPL;

    return m_tlib2->GetCustData(guid, pVarVal);
}

STDMETHODIMP FauxTypeLib2::GetLibStatistics(ULONG* pcUniqueNames, ULONG* pcchUniqueNames)
{
    if (!m_tlib2)
        return E_NOTIMPL;

    return m_tlib2->GetLibStatistics(pcUniqueNames, pcchUniqueNames);
}

STDMETHODIMP FauxTypeLib2::GetDocumentation2(INT index, LCID lcid, BSTR* pbstrHelpString, DWORD* pdwHelpStringContext, BSTR* pbstrHelpStringDll)
{
    if (!m_tlib2)
        return E_NOTIMPL;

    return m_tlib2->GetDocumentation2(index, lcid, pbstrHelpString, pdwHelpStringContext, pbstrHelpStringDll);
}

STDMETHODIMP FauxTypeLib2::GetAllCustData(CUSTDATA* pCustData)
{
    if (!m_tlib2)
        return E_NOTIMPL;

    return m_tlib2->GetAllCustData(pCustData);
}

