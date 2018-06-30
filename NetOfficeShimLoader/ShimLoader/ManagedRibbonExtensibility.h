#pragma once
#include "stdafx.h"
#include "IRibbonExtensibility.h"

extern HINSTANCE _module;
extern ULONG _components;
extern ULONG _locks;

class ManagedRibbonExtensibility : public IRibbonExtensibility
{

public:

	// Ctor, Dtor
	ManagedRibbonExtensibility(IUnknown* innerUnkown);
	~ManagedRibbonExtensibility();

	// IRibbonExtensibility Implementation
	STDMETHOD(GetCustomUI)(BSTR RibbonID, BSTR* RibbonXml);

	// IDispatch Implementation
	STDMETHODIMP GetTypeInfoCount(UINT* pctinfo);
	STDMETHODIMP GetTypeInfo(UINT iTInfo, LCID lcid, ITypeInfo** ppTInfo);
	STDMETHODIMP GetIDsOfNames(REFIID riid, LPOLESTR* rgszNames, UINT cNames, LCID lcid, DISPID* rgDispId);
	STDMETHODIMP Invoke(DISPID dispIdMember, REFIID riid, LCID lcid, WORD wFlags, DISPPARAMS* pDispParams, VARIANT* pVarResult, EXCEPINFO* pExcepInfo, UINT* puArgErr);

	// IUnknown Implementation
	STDMETHODIMP         QueryInterface(REFIID riid, void ** ppv);
	STDMETHODIMP_(ULONG) AddRef(void);
	STDMETHODIMP_(ULONG) Release(void);

private:

	IUnknown*					_innerUnkown;
	ULONG						_refCounter;

};
