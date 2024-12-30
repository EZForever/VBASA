#pragma once
#include "stdafx.h"

class FauxTypeLib2 final : public ITypeLib2
{
private:
    ULONG m_refcnt = 0;
    ITypeLib* m_tlib = NULL;
    ITypeLib2* m_tlib2 = NULL;

    FauxTypeLib2(ITypeLib* tlib, ITypeLib2* tlib2)
        : m_tlib(tlib), m_tlib2(tlib2) {}
    FauxTypeLib2(FauxTypeLib2&&) = delete;

    ~FauxTypeLib2() = default;

public:
    static ITypeLib2* CreateInstance(ITypeLib* tlib);

    // --- Implements ITypeLib2->ITypeLib->IUnknown

    STDMETHODIMP QueryInterface(REFIID riid, void** ppvObject);
    STDMETHODIMP_(ULONG) AddRef();
    STDMETHODIMP_(ULONG) Release();

    // --- Implements ITypeLib2->ITypeLib

    STDMETHODIMP_(UINT) GetTypeInfoCount();
    STDMETHODIMP GetTypeInfo(UINT index, ITypeInfo** ppTInfo);
    STDMETHODIMP GetTypeInfoType(UINT index, TYPEKIND* pTKind);
    STDMETHODIMP GetTypeInfoOfGuid(REFGUID guid, ITypeInfo** ppTinfo);
    STDMETHODIMP GetLibAttr(TLIBATTR** ppTLibAttr);
    STDMETHODIMP GetTypeComp(ITypeComp** ppTComp);
    STDMETHODIMP GetDocumentation(INT index, BSTR* pBstrName, BSTR* pBstrDocString, DWORD* pdwHelpContext, BSTR* pBstrHelpFile);
    STDMETHODIMP IsName(LPOLESTR szNameBuf, ULONG lHashVal, BOOL* pfName);
    STDMETHODIMP FindName(LPOLESTR szNameBuf, ULONG lHashVal, ITypeInfo** ppTInfo, MEMBERID* rgMemId, USHORT* pcFound);
    STDMETHODIMP_(void) ReleaseTLibAttr(TLIBATTR* pTLibAttr);

    // --- Implements ITypeLib2

    STDMETHODIMP GetCustData(REFGUID guid, VARIANT* pVarVal);
    STDMETHODIMP GetLibStatistics(ULONG* pcUniqueNames, ULONG* pcchUniqueNames);
    STDMETHODIMP GetDocumentation2(INT index, LCID lcid, BSTR* pbstrHelpString, DWORD* pdwHelpStringContext, BSTR* pbstrHelpStringDll);
    STDMETHODIMP GetAllCustData(CUSTDATA* pCustData);
};

