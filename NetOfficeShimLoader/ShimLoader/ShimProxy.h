#pragma once
#include "stdAfx.h"
#include "ClrHost.h"
#include "CLRUpdateHost.h"
#include "Aggregators.h"
#include "IShimProxy.hpp"
#include "Extensibility2.h"
#include "ManagedCustomTaskPaneConsumer.h"
#include "ShimHost.h"

extern HANDLE		_thread;
extern HINSTANCE	_module;
extern BOOL			_unloadAllowed;
extern void IncComponents(LPCWSTR type);
extern void DecComponents(LPCWSTR type);
extern BOOL ENABLE_SHIM;
extern BOOL ENABLE_DEBUG_MESSAGE_BOX;
extern HRESULT EXTENSIBILITY_DEFAULT_RESULT;
extern HRESULT EXTENSIBILITY_FAIL_RESULT;

namespace NetOffice_ShimLoader
{
	class ShimProxy : public IDTExtensibility2, public IShimProxy
	{

	public:

		// Ctor, Dtor
		ShimProxy();
		virtual ~ShimProxy();

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
		STDMETHODIMP LoadUpdateHandler();
		STDMETHODIMP Update(BOOL async);

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
		CLRUpdateHost*						_updateLoader;
		ShimHost*							_shimHost;
		ManagedRibbonExtensibility*			_ribbonExtensibility;
		ManagedCustomTaskPaneConsumer*		_paneConsumer;
		volatile DWORD						_currentReloadTread;
		IDispatch*							_application;
		IDispatch*							_addInInst;
		ext_ConnectMode						_connectMode;
		LPSAFEARRAY*						_customOnConnectionArgs;
		LPSAFEARRAY*						_customOnAddInsUpdateArgs;
		LPSAFEARRAY*						_customOnStartupCompleteArgs;
		LPSAFEARRAY*						_customEmptyArgs;
		BOOL								_onConnectionPassed;
		BOOL								_onAddInsUpdatePassed;
		BOOL								_onStartupCompletePassed;
	};
}
