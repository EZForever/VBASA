#pragma once
#include "stdafx.h"

class CommandBarControlEvents final : public VBIDE::_dispCommandBarControlEvents
{
private:
    ULONG m_refcnt = 0;
    ITypeInfo* m_tinfo = NULL;

    CommandBarControlEvents() = default;
    CommandBarControlEvents(CommandBarControlEvents&&) = delete;

    ~CommandBarControlEvents() = default;

public:
    static VBIDE::_dispCommandBarControlEvents* CreateInstance();

    // --- Implements _dispCommandBarControlEvents->IDispatch->IUnknown

    STDMETHODIMP QueryInterface(REFIID riid, void** ppvObject);
    STDMETHODIMP_(ULONG) AddRef();
    STDMETHODIMP_(ULONG) Release();

    // --- Implements _dispCommandBarControlEvents->IDispatch

    STDMETHODIMP GetTypeInfoCount(UINT* pctinfo);
    STDMETHODIMP GetTypeInfo(UINT iTInfo, LCID lcid, ITypeInfo** ppTInfo);
    STDMETHODIMP GetIDsOfNames(REFIID riid, LPOLESTR* rgszNames, UINT cNames, LCID lcid, DISPID* rgDispId);
    STDMETHODIMP Invoke(DISPID dispIdMember, REFIID riid, LCID lcid, WORD wFlags, DISPPARAMS* pDispParams, VARIANT* pVarResult, EXCEPINFO* pExcepInfo, UINT* puArgErr);

    // --- Implements _dispCommandBarControlEvents

    virtual STDMETHODIMP_(VOID) Click(IDispatch* CommandBarControl, VARIANT_BOOL* handled, VARIANT_BOOL* CancelDefault);
};

