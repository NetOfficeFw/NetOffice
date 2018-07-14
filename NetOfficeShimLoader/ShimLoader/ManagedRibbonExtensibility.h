#pragma once
#include "stdafx.h"
#include "IRibbonExtensibility.h"
#include "IShimProxy.hpp"
//#include "Vars.hpp"

extern HANDLE		_thread;
extern HINSTANCE	_module;
extern void IncComponents(LPCWSTR type);
extern void DecComponents(LPCWSTR type);

namespace NetOffice_ShimLoader
{
	class ManagedRibbonExtensibility : public IRibbonExtensibility
	{

	public:

		// Ctor, Dtor
		ManagedRibbonExtensibility(IShimProxy* parent, IRibbonExtensibility* innerExtensibility);
		virtual ~ManagedRibbonExtensibility();

		// ManagedRibbonExtensibility Methods
		STDMETHODIMP SetInnerPointer(IRibbonExtensibility* innerExtensibility);

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

		IShimProxy * _parent;
		IRibbonExtensibility*		_innerExtensibility;
		ULONG						_refCounter;

	};
}
