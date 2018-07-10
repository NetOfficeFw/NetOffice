#pragma once
#include "stdafx.h"
#include "Aggregators.h"
#include "Extensibility2.h"
#include "IRibbonExtensibility.h"
#include "ICTPFactory.h"

extern HANDLE		_thread;
extern HINSTANCE	_module;
extern void IncComponents(LPCWSTR type);
extern void DecComponents(LPCWSTR type);

namespace NetOffice_ShimLoader
{
	class ManagedAddin : public IDTExtensibility2
	{

	public:

		// Ctor, Dtor
		ManagedAddin(IUnknown* innerUnkown);
		virtual ~ManagedAddin();

		// ManagedAddin Methods
		IUnknown* InnerUnkown();
		STDMETHODIMP ReloadNotification(BSTR custom);

		// IDTExtensibility2 Implementation
		STDMETHODIMP OnConnection(IDispatch* application, ext_ConnectMode connectMode, IDispatch* addInInst, LPSAFEARRAY* custom);
		STDMETHODIMP OnDisconnection(ext_DisconnectMode removeMode, LPSAFEARRAY* custom);
		STDMETHODIMP OnAddInsUpdate(LPSAFEARRAY* custom);
		STDMETHODIMP OnStartupComplete(LPSAFEARRAY* custom);
		STDMETHODIMP OnBeginShutdown(LPSAFEARRAY* custom);

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

		IUnknown *	_innerUnkown;
		ULONG		_refCounter;

	};
}
