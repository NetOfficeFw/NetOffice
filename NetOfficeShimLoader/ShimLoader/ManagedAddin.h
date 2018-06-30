#pragma once
#include "stdafx.h"
#include "Aggregators.h"
#include "Extensibility2.h"
#include "IRibbonExtensibility.h"
#include "ICTPFactory.h"

extern HINSTANCE _module;
extern ULONG _components;
extern ULONG _locks;

class ManagedAddin : public IDTExtensibility2, public IRibbonExtensibility, public ICustomTaskPaneConsumer
{

public:

	// Ctor, Dtor
	ManagedAddin(IUnknown* innerUnkown);
	~ManagedAddin();

	// IDTExtensibility2 Implementation
	STDMETHODIMP OnConnection(IDispatch* application, ext_ConnectMode connectMode, IDispatch* addInInst, LPSAFEARRAY* custom);
	STDMETHODIMP OnDisconnection(ext_DisconnectMode removeMode, LPSAFEARRAY* custom);
	STDMETHODIMP OnAddInsUpdate(LPSAFEARRAY* custom);
	STDMETHODIMP OnStartupComplete(LPSAFEARRAY* custom);
	STDMETHODIMP OnBeginShutdown(LPSAFEARRAY* custom);

	// IRibbonExtensibility Implementation
	STDMETHOD(GetCustomUI)(BSTR RibbonID, BSTR* RibbonXml);

	// ICustomTaskPaneConsumer Implementation
	STDMETHOD(CTPFactoryAvailable) (ICTPFactory* CTPFactoryInst);

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
