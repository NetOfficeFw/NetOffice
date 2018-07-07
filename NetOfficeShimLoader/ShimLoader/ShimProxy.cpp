#include "StdAfx.h"
#include "ShimProxy.h"

DWORD WINAPI ReloadCLRInternal(LPVOID lpParameter);

/***************************************************************************
* Ctor Dtor
***************************************************************************/

ShimProxy::ShimProxy()
{
	_refCounter = 0;
	_updateAggregator = nullptr;
	_loader = nullptr;
	_ribbonExtensibility = nullptr;
	_paneConsumer = nullptr;
	_currentReloadTread = nullptr;
	_application = nullptr;
	_addInInst = nullptr;
	_customOnConnection = nullptr;
	_customOnAddInsUpdate = nullptr;
	_customOnStartupComplete = nullptr;
	_customEmptyArgs = new LPSAFEARRAY();
	_connectMode = static_cast<ext_ConnectMode>(0);
	_components++;

	// load the CLR here because host application may call QueryInterface before OnConnection
	if (ENABLE_SHIM)
	{
		_updateAggregator = new OuterUpdateAggregator(this);
		IOuterUpdateAggregator* updateAggregator = static_cast<IOuterUpdateAggregator*>(_updateAggregator);
		_loader = new (std::nothrow) ClrHost(updateAggregator);

		//DWORD dwThreadId;
		//HANDLE handle =_currentReloadTread = CreateThread(
		//	(SECURITY_ATTRIBUTES *)0,
		//	0,
		//	&ReloadCLRInternal,
		//	this,
		//	0,
		//	&dwThreadId
		//);

		//WaitForSingleObject(handle, INFINITE);

		_loader->Load();
	}
}

ShimProxy::~ShimProxy()
{
	Cleanup();
	_components--;
}


/***************************************************************************
* ShimProxy Methods
***************************************************************************/

STDMETHODIMP ShimProxy::Cleanup()
{
	HRESULT hr = S_OK;

	CloseReloadThread();

	_customOnConnection = nullptr;
	_customOnAddInsUpdate = nullptr;
	_customOnStartupComplete = nullptr;
	if (_customEmptyArgs)
	{
		_customEmptyArgs = nullptr;
		delete _customEmptyArgs;
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
	if (_updateAggregator)
	{
		delete _updateAggregator;
		_updateAggregator = nullptr;
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
		_customOnConnection = custom;
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

	ShimDebugMessageBox(L"Fail", L"OnConnection");
	ValidateExtensibilityFail(hr);
	return hr;
}

STDMETHODIMP ShimProxy::OnDisconnection(ext_DisconnectMode removeMode, LPSAFEARRAY* custom)
{
	HRESULT hr = EXTENSIBILITY_DEFAULT_RESULT;

	try
	{
		if (ENABLE_SHIM && IsCLRLoaded())
		{
			IfFailGo(_loader->OuterAggregator()->Addin()->OnDisconnection(removeMode, custom));
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
		_customOnAddInsUpdate = custom;
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
		_customOnStartupComplete = custom;
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
	return NULL != _currentReloadTread ? TRUE : FALSE;
}

BOOL ShimProxy::IsAsyncReloadThreadInProgress()
{
	if (_currentReloadTread)
	{
		return S_OK != _currentReloadTread ? TRUE : FALSE;
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
		//DWORD dwThreadId;
		//_currentReloadTread = CreateThread(
		//	(SECURITY_ATTRIBUTES*)0,
		//	0,
		//	&ReloadCLRInternal,
		//	this,
		//	0,
		//	&dwThreadId
		//);
		// execute in ui thread
		QueueUserAPC((PAPCFUNC)ReloadCLRInternal, _thread, (ULONG_PTR)this);
	}
	else
	{
		_currentReloadTread = S_OK;
		hr = ReloadCLRInternal(this);
	}

	return hr;
}

STDMETHODIMP ShimProxy::CloseReloadThread()
{
	HRESULT hr = S_OK;
	if (NULL != _currentReloadTread)
	{
		if (S_OK != _currentReloadTread)
		{
			if (!CloseHandle(_currentReloadTread))
				hr = E_FAIL;
		}
		_currentReloadTread = nullptr;
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

		IfFailGo(_loader->OuterAggregator()->Addin()->OnConnection(_application, _connectMode, _addInInst, _customEmptyArgs));
		IfFailGo(_loader->OuterAggregator()->Addin()->OnAddInsUpdate(_customEmptyArgs));
		if (_paneConsumer)
		{
			res = _paneConsumer->CTPFactoryAvailable(_paneConsumer->InnerCtpFactory());
			if (!SUCCEEDED(res))
				hr = res;
		}
		IfFailGo(_loader->OuterAggregator()->Addin()->OnStartupComplete(_customEmptyArgs));
	}
	else
	{
		hr = E_UNEXPECTED;
	}

	return hr;

Error:

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
			if(SUCCEEDED(hr))
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
	if(0 == _refCounter)
		delete this;
	return 0;
}
