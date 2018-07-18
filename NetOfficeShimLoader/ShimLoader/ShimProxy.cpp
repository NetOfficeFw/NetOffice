#include "StdAfx.h"
#include "ShimProxy.h"
#include "Vars.h"

#include "ShimArguments.h"

namespace NetOffice_ShimLoader
{
	DWORD WINAPI ReloadCLRInternal(LPVOID lpParameter);
	DWORD WINAPI UpdateInternal(LPVOID lpParameter);

	/***************************************************************************
	* Ctor Dtor
	***************************************************************************/

	ShimProxy::ShimProxy()
	{
		_refCounter = 0;
		_shimHost = nullptr;
		_loader = nullptr;
		_updateLoader = nullptr;
		_ribbonExtensibility = nullptr;
		_paneConsumer = nullptr;
		_currentReloadTread = 0;
		_application = nullptr;
		_addInInst = nullptr;
		_customOnConnectionArgs = nullptr;
		_customOnAddInsUpdateArgs = nullptr;
		_customOnStartupCompleteArgs = nullptr;
		_onConnectionPassed = FALSE;
		_onAddInsUpdatePassed = FALSE;
		_onStartupCompletePassed = FALSE;
		_customEmptyArgs = new LPSAFEARRAY();
		_connectMode = static_cast<ext_ConnectMode>(0);

		IncComponents(L"ShimProxy");

		// load the CLR here because host application may call QueryInterface before OnConnection
		if (ENABLE_SHIM)
		{
			_shimHost = new ShimHost(this);
			IShimHost* shimHost = static_cast<IShimHost*>(_shimHost);
			_loader = new (std::nothrow) ClrHost(shimHost);
			if (_loader)
			{
				if(!SUCCEEDED(_loader->Load()))
					DebugOutput(L"ClrHost::Load failed.");
			}
			else
			{
				DebugOutput(L"ClrHost::ClrHost failed.");
			}
		}
	}

	ShimProxy::~ShimProxy()
	{
		if (!SUCCEEDED(Cleanup()))
			DebugOutput(L"ShimProxy::Cleanup failed.");
		DecComponents(L"ShimProxy");
		_unloadAllowed = TRUE;
	}


	/***************************************************************************
	* ShimProxy Methods
	***************************************************************************/

	STDMETHODIMP ShimProxy::Cleanup()
	{
		HRESULT hr = S_OK;

		CloseReloadThread();

		if (_customOnConnectionArgs)
		{
			HRESULT sad = SafeArrayDestroy(*_customOnConnectionArgs);
			if (!SUCCEEDED(sad))
				hr = sad;
			delete _customOnConnectionArgs;
			_customOnConnectionArgs = nullptr;
		}
		if (_customOnAddInsUpdateArgs)
		{
			HRESULT sad = SafeArrayDestroy(*_customOnAddInsUpdateArgs);
			if (!SUCCEEDED(sad))
				hr = sad;
			delete _customOnAddInsUpdateArgs;
			_customOnAddInsUpdateArgs = nullptr;
		}
		if (_customOnStartupCompleteArgs)
		{
			HRESULT sad = SafeArrayDestroy(*_customOnStartupCompleteArgs);
			if (!SUCCEEDED(sad))
				hr = sad;
			delete _customOnStartupCompleteArgs;
			_customOnStartupCompleteArgs = nullptr;
		}
		if (_customEmptyArgs)
		{
			delete _customEmptyArgs;
			_customEmptyArgs = nullptr;
		}
		if (_paneConsumer)
		{
			delete _paneConsumer;
			_paneConsumer = nullptr;
		}
		if (_ribbonExtensibility)
		{
			delete _ribbonExtensibility;
			_ribbonExtensibility = nullptr;
		}
		if (_loader)
		{
			delete _loader;
			_loader = nullptr;
		}
		if (_shimHost)
		{
			delete _shimHost;
			_shimHost = nullptr;
		}
		if (_updateLoader)
		{
			delete _updateLoader;
			_updateLoader = nullptr;
		}
		return hr;
	}


	/***************************************************************************
	* IDTExtensibility2 Implementation
	***************************************************************************/

	STDMETHODIMP ShimProxy::OnConnection(IDispatch* application, ext_ConnectMode connectMode, IDispatch* addInInst, LPSAFEARRAY* custom)
	{
		HRESULT hr = EXTENSIBILITY_DEFAULT_RESULT;

		try
		{
			_onConnectionPassed = TRUE;
			_unloadAllowed = FALSE;

			if (custom && (*custom))
			{
				_customOnConnectionArgs = new LPSAFEARRAY();
				if (!SUCCEEDED(SafeArrayCopy((*custom), _customOnConnectionArgs)))
				{
					delete _customOnConnectionArgs;
					_customOnConnectionArgs = nullptr;
				}
			}
			if (application)
			{
				_application = application;
				_application->AddRef();

			}
			if (addInInst)
			{
				_addInInst = addInInst;
				_addInInst->AddRef();
			}

			if (ENABLE_SHIM && IsCLRLoaded())
			{
				IfFailGo(_loader->OuterAggregator()->Addin()->OnConnection(application, connectMode, addInInst, custom));
			}
			else if (_loader)
			{
				delete _loader;
				_loader = nullptr;
			}
			if (ENABLE_SHIM && !IsCLRLoaded())
			{
				hr = E_FAIL;
				goto Error;
			}

			return hr;
		}
		catch (...)
		{
			ShimDebugMessageBox(L"Error", L"OnConnection");
			hr = E_FAIL;
		}

	Error:

		_unloadAllowed = TRUE;
		ShimDebugMessageBox(L"Fail", L"OnConnection");
		ValidateExtensibilityFail(hr);
		return hr;
	}

	STDMETHODIMP ShimProxy::OnDisconnection(ext_DisconnectMode removeMode, LPSAFEARRAY* custom)
	{
		HRESULT hr = EXTENSIBILITY_DEFAULT_RESULT;

		try
		{
			_unloadAllowed = TRUE;

			if (ENABLE_SHIM && IsCLRLoaded())
			{
				IfFailGo(_loader->OuterAggregator()->Addin()->OnDisconnection(removeMode, custom));
			}

			if (_customOnConnectionArgs)
			{
				SafeArrayDestroy(*_customOnConnectionArgs);
				delete _customOnConnectionArgs;
				_customOnConnectionArgs = nullptr;
			}
			if (_customOnAddInsUpdateArgs)
			{
				SafeArrayDestroy(*_customOnAddInsUpdateArgs);
				delete _customOnAddInsUpdateArgs;
				_customOnAddInsUpdateArgs = nullptr;
			}
			if (_customOnStartupCompleteArgs)
			{
				SafeArrayDestroy(*_customOnStartupCompleteArgs);
				delete _customOnStartupCompleteArgs;
				_customOnStartupCompleteArgs = nullptr;
			}

			// no cleanup call here because host application may still holds IUnkown Pointer to ribbon/taskpane/etc.

			if (_application)
			{
				_application->Release();
				_application = nullptr;
			}
			if (_addInInst)
			{
				_addInInst->Release();
				_addInInst = nullptr;
			}

			return hr;
		}
		catch (...)
		{
			ShimDebugMessageBox(L"Error", L"OnDisconnection");
			hr = E_FAIL;
		}

	Error:

		ShimDebugMessageBox(L"Fail", L"OnDisconnection");
		ValidateExtensibilityFail(hr);
		return hr;
	}

	STDMETHODIMP ShimProxy::OnAddInsUpdate(LPSAFEARRAY* custom)
	{
		HRESULT hr = EXTENSIBILITY_DEFAULT_RESULT;

		try
		{
			_onAddInsUpdatePassed = TRUE;

			_customOnAddInsUpdateArgs = new LPSAFEARRAY();
			if (!SUCCEEDED(SafeArrayCopy((*custom), _customOnAddInsUpdateArgs)))
			{
				delete _customOnAddInsUpdateArgs;
				_customOnAddInsUpdateArgs = nullptr;
			}

			if (ENABLE_SHIM && IsCLRLoaded())
			{
				IfFailGo(_loader->OuterAggregator()->Addin()->OnAddInsUpdate(custom));
			}
		}
		catch (...)
		{
			ShimDebugMessageBox(L"Error", L"OnAddInsUpdate");
			hr = E_FAIL;
		}

		return hr;

	Error:

		ShimDebugMessageBox(L"Fail", L"OnAddInsUpdate");
		ValidateExtensibilityFail(hr);
		return hr;
	}

	STDMETHODIMP ShimProxy::OnStartupComplete(LPSAFEARRAY* custom)
	{
		HRESULT hr = EXTENSIBILITY_DEFAULT_RESULT;

		try
		{
			_onStartupCompletePassed = TRUE;

			_customOnStartupCompleteArgs = new LPSAFEARRAY();
			if (!SUCCEEDED(SafeArrayCopy((*custom), _customOnStartupCompleteArgs)))
			{
				delete _customOnStartupCompleteArgs;
				_customOnStartupCompleteArgs = nullptr;
			}

			if (ENABLE_SHIM && IsCLRLoaded())
			{
				IfFailGo(_loader->OuterAggregator()->Addin()->OnStartupComplete(custom));
			}
		}
		catch (...)
		{
			MessageBox(GetDesktopWindow(), L"Error", L"OnStartupComplete", 0);
			hr = E_FAIL;
		}

		return hr;

	Error:

		ShimDebugMessageBox(L"Fail", L"OnStartupComplete");
		ValidateExtensibilityFail(hr);
		return hr;
	}

	STDMETHODIMP ShimProxy::OnBeginShutdown(LPSAFEARRAY* custom)
	{
		HRESULT hr = EXTENSIBILITY_DEFAULT_RESULT;

		try
		{
			if (ENABLE_SHIM && IsCLRLoaded())
			{
				IfFailGo(_loader->OuterAggregator()->Addin()->OnBeginShutdown(custom));
			}

		}
		catch (...)
		{
			ShimDebugMessageBox(L"Error", L"OnBeginShutdown");
			hr = E_FAIL;
		}

		return hr;

	Error:

		ShimDebugMessageBox(L"Fail", L"OnBeginShutdown");
		ValidateExtensibilityFail(hr);
		return hr;
	}


	/***************************************************************************
	* IShimProxy Implementation
	***************************************************************************/

	BOOL STDMETHODCALLTYPE ShimProxy::IsCLRLoaded()
	{
		BOOL result = FALSE;
		result = _loader && _loader->IsLoaded();
		return result;
	}

	STDMETHODIMP ShimProxy::LoadCLR()
	{
		HRESULT hr = E_FAIL;
		if (!IsCLRLoaded())
		{
			IfFailGo(_loader->Load());
		}

		return hr;

	Error:

		return hr;
	}

	STDMETHODIMP ShimProxy::UnloadCLR()
	{
		HRESULT hr = E_FAIL;
		if (IsCLRLoaded())
		{
			IfFailGo(_loader->Unload());
		}

		return hr;
	Error:
		return hr;
	}

	BOOL ShimProxy::IsReloadThreadInProgress()
	{
		return 1 == _currentReloadTread ? TRUE : FALSE;
	}

	BOOL ShimProxy::IsAsyncReloadThreadInProgress()
	{
		if (_currentReloadTread)
		{
			return 2 == _currentReloadTread ? TRUE : FALSE;
		}
		else
		{
			return FALSE;
		}
	}

	STDMETHODIMP ShimProxy::ReloadCLR(BOOL async)
	{
		if (IsReloadThreadInProgress())
			return E_UNEXPECTED;

		HRESULT hr = S_OK;

		if (async)
		{
			if (0 == QueueUserAPC((PAPCFUNC)ReloadCLRInternal, _thread, (ULONG_PTR)this))
				hr = E_FAIL;
			_currentReloadTread = 2;
		}
		else
		{
			_currentReloadTread = 1;
			hr = ReloadCLRInternal(this);
		}

		return hr;
	}

	STDMETHODIMP ShimProxy::CloseReloadThread()
	{
		HRESULT hr = S_OK;
		if (0 != _currentReloadTread)
		{
			_currentReloadTread = 0;
		}
		else
		{
			hr = E_UNEXPECTED;
		}
		return hr;
	}

	STDMETHODIMP ShimProxy::AssignInnerPointers()
	{
		HRESULT hr = S_OK;

		if (IsCLRLoaded())
		{
			HRESULT res = S_OK;

			if (_paneConsumer)
			{
				IUnknown* unknown = _loader->OuterAggregator()->Addin()->InnerUnkown();
				ICustomTaskPaneConsumer* consumer = nullptr;
				if (SUCCEEDED(unknown->QueryInterface(IID_ICustomTaskPaneConsumer, (LPVOID*)&consumer)))
				{
					HRESULT setResult = _paneConsumer->SetInnerPointer(consumer);
					if (!SUCCEEDED(setResult))
						hr = setResult;
				}
				else
				{
					hr = E_FAIL;
				}
			}

			if (_ribbonExtensibility)
			{
				IUnknown* unknown = _loader->OuterAggregator()->Addin()->InnerUnkown();
				IRibbonExtensibility* ribbon = nullptr;
				if (SUCCEEDED(unknown->QueryInterface(IID_IRibbonExtensibility, (LPVOID*)&ribbon)))
				{
					HRESULT setResult = _ribbonExtensibility->SetInnerPointer(ribbon);
					if (!SUCCEEDED(setResult))
						hr = setResult;
				}
				else
				{
					hr = E_FAIL;
				}
			}

			BSTR customData = _shimHost->CustomData();
			_loader->OuterAggregator()->Addin()->ReloadNotification(customData);

			if(_onConnectionPassed)
				IfFailGo(_loader->OuterAggregator()->Addin()->OnConnection(_application, _connectMode, _addInInst,  NULL != _customOnConnectionArgs ? _customOnConnectionArgs : _customEmptyArgs));
			if(_onAddInsUpdatePassed)
				IfFailGo(_loader->OuterAggregator()->Addin()->OnAddInsUpdate(NULL != _customOnAddInsUpdateArgs ? _customOnAddInsUpdateArgs : _customEmptyArgs));
			if (_paneConsumer)
			{
				res = _paneConsumer->CTPFactoryAvailable(_paneConsumer->InnerCtpFactory());
				if (!SUCCEEDED(res))
					hr = res;
			}
			if(_onStartupCompletePassed)
				IfFailGo(_loader->OuterAggregator()->Addin()->OnStartupComplete(NULL != _customOnStartupCompleteArgs ? _customOnStartupCompleteArgs : _customEmptyArgs));
		}
		else
		{
			hr = E_UNEXPECTED;
		}

		return hr;

	Error:

		return hr;
	}

	STDMETHODIMP ShimProxy::LoadUpdateHandler()
	{
		HRESULT hr = E_FAIL;
		if (!IsCLRLoaded() && NULL == _updateLoader)
		{
			_updateLoader = new (std::nothrow) CLRUpdateHost();
			if (_updateLoader)
				hr = _updateLoader->Load();
			else
				hr = E_OUTOFMEMORY;

			if (SUCCEEDED(hr))
			{
				ManagedUpdateHandler* updater = _updateLoader->OuterAggregator()->ManagedUpdater();
				if (updater && _application)
				{
					hr = updater->SetApplication(_application);
				}

				if (updater && SUCCEEDED(hr))
				{
					BSTR customData = _shimHost->CustomData();
					hr = updater->SetCustomData(customData);
				}

				if (updater && SUCCEEDED(hr) && updater->CanExecute())
				{
					updater->Execute();
				}

				if (updater)
					updater->Close();

				BSTR customData = _updateLoader->Host()->CustomData();
				if(customData)
					_shimHost->SetCustomData(customData);

				hr = _updateLoader->Unload();

				if (SUCCEEDED(hr))
					hr = ReloadCLR(FALSE);
			}
		}
		else
		{
			hr = E_UNEXPECTED;
			goto Error;
		}

		if (_updateLoader)
		{
			delete _updateLoader;
			_updateLoader = nullptr;
		}
		return hr;

	Error:

		if (_updateLoader)
		{
			delete _updateLoader;
			_updateLoader = nullptr;
		}
		ReloadCLR(false);
		return hr;
	}

	STDMETHODIMP ShimProxy::Update(BOOL async)
	{
		if (IsReloadThreadInProgress())
			return E_UNEXPECTED;

		HRESULT hr = S_OK;

		if (async)
		{
			if (0 == QueueUserAPC((PAPCFUNC)UpdateInternal, _thread, (ULONG_PTR)this))
				hr = E_FAIL;
			_currentReloadTread = 2;
		}
		else
		{
			_currentReloadTread = 1;
			hr = UpdateInternal(this);
		}

		return hr;
	}

	DWORD WINAPI UpdateInternal(LPVOID lpParameter)
	{
		HRESULT hr = S_OK;
		ShimProxy* proxy = (ShimProxy*)lpParameter;

		if (proxy->IsCLRLoaded())
		{
			hr = proxy->UnloadCLR();
		}

		if (!proxy->IsCLRLoaded())
		{
			hr = proxy->LoadUpdateHandler();
		}

		return hr;
	}

	DWORD WINAPI ReloadCLRInternal(LPVOID lpParameter)
	{
		HRESULT hr = S_OK;
		ShimProxy* proxy = (ShimProxy*)lpParameter;

		if (proxy->IsCLRLoaded())
		{
			hr = proxy->UnloadCLR();
		}

		if (!proxy->IsCLRLoaded())
		{
			hr = proxy->LoadCLR();
		}

		if (SUCCEEDED(hr))
			hr = proxy->AssignInnerPointers();

		proxy->CloseReloadThread();

		return hr;
	}


	/***************************************************************************
	* IDispatch Implementation
	***************************************************************************/

	STDMETHODIMP ShimProxy::GetTypeInfoCount(UINT* pctinfo)
	{
		return E_FAIL;
	}

	STDMETHODIMP ShimProxy::GetTypeInfo(UINT iTInfo, LCID lcid, ITypeInfo** ppTInfo)
	{
		return E_FAIL;
	}

	STDMETHODIMP ShimProxy::GetIDsOfNames(REFIID riid, LPOLESTR* rgszNames, UINT cNames, LCID lcid, DISPID* rgDispId)
	{
		return E_FAIL;
	}

	STDMETHODIMP ShimProxy::Invoke(DISPID dispIdMember, REFIID riid, LCID lcid, WORD wFlags, DISPPARAMS* pDispParams, VARIANT* pVarResult, EXCEPINFO* pExcepInfo, UINT* puArgErr)
	{
		return E_FAIL;
	}


	/***************************************************************************
	* IUnknown Implementation
	***************************************************************************/

	STDMETHODIMP ShimProxy::QueryInterface(REFIID riid, void** ppv)
	{
		if (NULL == ppv)
			return E_POINTER;
		*ppv = NULL;

		HRESULT hr = E_FAIL;
		bool isBlind = false;

		if (((IID_IDTExtensibility2 == riid) || (IID_IUnknown == riid) || (IID_IDispatch == riid)))
		{
			*ppv = static_cast<IDTExtensibility2*>(this);
			hr = S_OK;
		}
		else if (ENABLE_SHIM && (IID_IRibbonExtensibility == riid) && IsCLRLoaded())
		{
			// wrap and cache to encapsulate from host application
			// this is to prevent conflicts when reload the CLR on the fly
			if (!_ribbonExtensibility)
			{
				IUnknown* inner = _loader->OuterAggregator()->Addin()->InnerUnkown();
				IRibbonExtensibility* ribbon = nullptr;
				hr = inner->QueryInterface(riid, (LPVOID*)&ribbon);
				if (SUCCEEDED(hr))
					_ribbonExtensibility = new (std::nothrow) ManagedRibbonExtensibility(static_cast<IShimProxy*>(this), ribbon);
				if (!_ribbonExtensibility && NULL != ribbon)
				{
					ribbon->Release();
					hr = E_OUTOFMEMORY;
				}
			}
			if (_ribbonExtensibility)
			{
				*ppv = static_cast<IRibbonExtensibility*>(_ribbonExtensibility);
				hr = S_OK;
			}
		}
		else if (ENABLE_SHIM && (IID_ICustomTaskPaneConsumer == riid) && IsCLRLoaded())
		{
			// wrap and cache to encapsulate from host application
			// this is to prevent conflicts when reload the CLR on the fly
			if (!_paneConsumer)
			{
				IUnknown* inner = _loader->OuterAggregator()->Addin()->InnerUnkown();
				ICustomTaskPaneConsumer* consumer = nullptr;
				hr = inner->QueryInterface(riid, (LPVOID*)&consumer);
				if (SUCCEEDED(hr))
					_paneConsumer = new (std::nothrow) ManagedCustomTaskPaneConsumer(static_cast<IShimProxy*>(this), consumer);
				if (!_paneConsumer && NULL != consumer)
				{
					consumer->Release();
					hr = E_OUTOFMEMORY;
				}
			}
			if (_paneConsumer)
			{
				*ppv = static_cast<ICustomTaskPaneConsumer*>(_paneConsumer);
				hr = S_OK;
			}
		}
		else if (ENABLE_SHIM && (!ENABLE_OUTER_UPDATE_AGGREGATOR && ENABLE_BLIND_AGGREGATION && IsCLRLoaded()))
		{
			// blind aggregation means the inner pointer is not bridged by the shim
			// so we can not reload the CLR on the fly because host application is not aware of the
			// fact that the pointers are no longer valid
			IUnknown* inner = _loader->OuterAggregator()->Addin()->InnerUnkown();
			hr = inner->QueryInterface(riid, ppv);
			isBlind = true;
		}

		if (NULL != *ppv && !isBlind)
		{
			reinterpret_cast<IUnknown*>(*ppv)->AddRef();
		}
		else if (NULL == *ppv)
		{
			hr = E_NOINTERFACE;
		}

		return hr;
	}

	STDMETHODIMP_(ULONG) ShimProxy::AddRef(void)
	{
		_refCounter++;
		return _refCounter;
	}

	STDMETHODIMP_(ULONG) ShimProxy::Release(void)
	{
		_refCounter--;
		if (0 == _refCounter)
			delete this;
		return 0;
	}
}
