#include "StdAfx.h"
#include "ShimProxy.h"


/***************************************************************************
* Ctor Dtor
***************************************************************************/

ShimProxy::ShimProxy()
{
	_refCounter = 0;
	_loader = nullptr;
	_ribbonExtensibility = nullptr;
	_paneConsumer = nullptr;
	_components++;

	// load the CLR here because host application may call QueryInterface before OnConnection
	_updateAggregator = new OuterUpdateAggregator(this);
	_loader = new (std::nothrow) ClrHost(_updateAggregator);
	_loader->Load();
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
	HRESULT hr = E_FAIL;
	if (IsCLRLoaded())
	{
		IfFailGo(_loader->OuterAggregator()->Addin()->OnConnection(application, connectMode, addInInst, custom));
	}
	else if(_loader)
	{
		delete _loader;
		_loader = nullptr;
		hr = S_OK; // host application want change load behavior from 3 to 2 if we signalize fail
	}

	return hr;
Error:
	return hr;
}

STDMETHODIMP ShimProxy::OnDisconnection(ext_DisconnectMode removeMode, LPSAFEARRAY* custom)
{
	HRESULT hr = E_FAIL;
	if (IsCLRLoaded())
	{
		IfFailGo(_loader->OuterAggregator()->Addin()->OnDisconnection(removeMode, custom));
	}
	// no cleanup call here because host application may still holds IUnkown Pointer to ribbon/taskpane/etc.

	return hr;
Error:
	return hr;
}

STDMETHODIMP ShimProxy::OnAddInsUpdate(LPSAFEARRAY* custom)
{
	HRESULT hr = E_FAIL;
	if (_loader && _loader->IsLoaded())
	{
		IfFailGo(_loader->OuterAggregator()->Addin()->OnAddInsUpdate(custom));
	}
	return hr;
Error:
	return hr;
}

STDMETHODIMP ShimProxy::OnStartupComplete(LPSAFEARRAY* custom)
{
	HRESULT hr = E_FAIL;
	if (IsCLRLoaded())
	{
		IfFailGo(_loader->OuterAggregator()->Addin()->OnStartupComplete(custom));
	}
	return hr;
Error:
	return hr;
}

STDMETHODIMP ShimProxy::OnBeginShutdown(LPSAFEARRAY* custom)
{
	HRESULT hr = E_FAIL;
	if (IsCLRLoaded())
	{
		IfFailGo(_loader->OuterAggregator()->Addin()->OnBeginShutdown(custom));
	}

	return hr;
Error:
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

STDMETHODIMP ShimProxy::ReloadCLR()
{
	HRESULT hr = E_FAIL;
	if (!IsCLRLoaded() && _loader)
	{
		IfFailGo(_loader->Load());

		// todo: store args in OnConnection and recall OnConnection/StartupComplete here

		if (_paneConsumer)
		{
			IUnknown* unknown = _loader->OuterAggregator()->Addin()->InnerUnkown();
			ICustomTaskPaneConsumer* consumer = nullptr;
			if (SUCCEEDED(unknown->QueryInterface(IID_IRibbonExtensibility, (LPVOID*)&consumer)))
				hr = _paneConsumer->SetInnerPointer(consumer);
		}

		if (_ribbonExtensibility)
		{
			IUnknown* unknown = _loader->OuterAggregator()->Addin()->InnerUnkown();
			IRibbonExtensibility* ribbon = nullptr;
			if (SUCCEEDED(unknown->QueryInterface(IID_IRibbonExtensibility, (LPVOID*)&ribbon)))
				hr = _ribbonExtensibility->SetInnerPointer(ribbon);
		}
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
	else if ((__uuidof(IRibbonExtensibility) == riid) && (_loader && _loader->IsLoaded()))
	{
		IUnknown* inner = _loader->OuterAggregator()->Addin()->InnerUnkown();
		IRibbonExtensibility* ribbon = nullptr;
		hr = inner->QueryInterface(riid, (LPVOID*)&ribbon);
		if (S_OK == hr)
		{
			if (!_ribbonExtensibility)
			{
				_ribbonExtensibility = new (std::nothrow) ManagedRibbonExtensibility(ribbon);
			}
			if (_ribbonExtensibility)
			{
				*ppv = static_cast<IRibbonExtensibility*>(_ribbonExtensibility);
				hr = S_OK;
			}
			else
			{
				ribbon->Release();
				hr = E_OUTOFMEMORY;
			}
		}
	}
	else if ((__uuidof(ICustomTaskPaneConsumer) == riid) && (_loader && _loader->IsLoaded()))
	{
		IUnknown* inner = _loader->OuterAggregator()->Addin()->InnerUnkown();
		ICustomTaskPaneConsumer* consumer = nullptr;
		hr = inner->QueryInterface(riid, (LPVOID*)&consumer);
		if (S_OK == hr)
		{
			if (!_paneConsumer)
			{
				_paneConsumer = new (std::nothrow) ManagedCustomTaskPaneConsumer(consumer);
			}
			if (_paneConsumer)
			{
				*ppv = static_cast<ICustomTaskPaneConsumer*>(_paneConsumer);
				hr = S_OK;
			}
			else
			{
				consumer->Release();
				hr = E_OUTOFMEMORY;
			}
		}
	}
	else if (!ENABLE_OUTER_UPDATE_AGGREGATOR && ENABLE_BLIND_AGGREGATION && _loader && _loader->IsLoaded())
	{
		IUnknown* inner = _loader->OuterAggregator()->Addin()->InnerUnkown();
		hr = inner->QueryInterface(riid, ppv);
		isBlind = true;
	}
	else
	{
		hr = E_NOINTERFACE;
	}

	if (NULL != *ppv && !isBlind)
	{
		reinterpret_cast<IUnknown*>(*ppv)->AddRef();
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
