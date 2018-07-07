#pragma once
#include "stdAfx.h"
#include "Vars.hpp"
#include "ClrHost.h"
#include "Aggregators.h"
#include "IShimProxy.hpp"
#include "Extensibility2.h"
#include "ManagedCustomTaskPaneConsumer.h"
#include "OuterUpdateAggregator.h"

extern HANDLE _thread;
extern HINSTANCE _module;
extern ULONG _components;
extern ULONG _locks;

class ShimProxy : public IDTExtensibility2, public IShimProxy
{

public:

	// Ctor, Dtor
	ShimProxy();
	~ShimProxy();

	// ShimProxy Methods
	STDMETHODIMP Cleanup();

	// IDTExtensibility2 Implementation
	STDMETHODIMP OnConnection(IDispatch* application, ext_ConnectMode connectMode, IDispatch* addInInst, LPSAFEARRAY* custom);
	STDMETHODIMP OnDisconnection(ext_DisconnectMode removeMode, LPSAFEARRAY* custom);
	STDMETHODIMP OnAddInsUpdate(LPSAFEARRAY* custom);
	STDMETHODIMP OnStartupComplete(LPSAFEARRAY* custom);
	STDMETHODIMP OnBeginShutdown(LPSAFEARRAY* custom);

	// IShimProxy Implementation
	BOOL STDMETHODCALLTYPE IsCLRLoaded();
	STDMETHODIMP LoadCLR();
	STDMETHODIMP UnloadCLR();
	BOOL IsReloadThreadInProgress();
	BOOL IsAsyncReloadThreadInProgress();
	STDMETHODIMP ReloadCLR(BOOL async);
	STDMETHODIMP CloseReloadThread();
	STDMETHODIMP AssignInnerPointers();

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

	ULONG								_refCounter;
	ClrHost*							_loader;
	OuterUpdateAggregator*				_updateAggregator;
	ManagedRibbonExtensibility*			_ribbonExtensibility;
	ManagedCustomTaskPaneConsumer*		_paneConsumer;
	volatile HANDLE						_currentReloadTread;
	IDispatch*							_application;
	IDispatch*							_addInInst;
	ext_ConnectMode						_connectMode;
	LPSAFEARRAY*						_customOnConnection;
	LPSAFEARRAY*						_customOnAddInsUpdate;
	LPSAFEARRAY*						_customOnStartupComplete;

	LPSAFEARRAY*						_customEmptyArgs;

};
