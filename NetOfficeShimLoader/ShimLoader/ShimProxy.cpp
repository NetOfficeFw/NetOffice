#include "StdAfx.h"
#include "ShimProxy.h"


/***************************************************************************
* Ctor Dtor
***************************************************************************/

ShimProxy::ShimProxy()
{
	_refCounter = 0;
	_loader = nullptr;
	_components++;

	_loader = new (std::nothrow) ClrHost();
	if (_loader)
		_loader->Load();
}

ShimProxy::~ShimProxy()
{
	if (_loader)
	{
		delete _loader;
		_loader = nullptr;
	}
	_components--;
}


/***************************************************************************
* IDTExtensibility2 Implementation
***************************************************************************/

STDMETHODIMP ShimProxy::OnConnection(IDispatch* application, ext_ConnectMode connectMode, IDispatch* addInInst, LPSAFEARRAY* custom)
{
	HRESULT hr = E_FAIL;
	if (_loader && _loader->IsLoaded())
	{
		hr = _loader->OuterAggregator()->Addin()->OnConnection(application, connectMode, addInInst, custom);
	}
	else if(_loader)
	{
		delete _loader;
		_loader = nullptr;
	}
	return hr;
}

STDMETHODIMP ShimProxy::OnDisconnection(ext_DisconnectMode removeMode, LPSAFEARRAY* custom)
{
	HRESULT hr = E_FAIL;
	if (_loader && _loader->IsLoaded())
	{
		hr = _loader->OuterAggregator()->Addin()->OnDisconnection(removeMode, custom);
	}
	if (_loader)
	{
		delete _loader;
		_loader = nullptr;
	}
	return hr;
}

STDMETHODIMP ShimProxy::OnAddInsUpdate(LPSAFEARRAY* custom)
{
	HRESULT hr = E_FAIL;
	if (_loader && _loader->IsLoaded())
	{
		hr = _loader->OuterAggregator()->Addin()->OnAddInsUpdate(custom);
	}
	return hr;
}

STDMETHODIMP ShimProxy::OnStartupComplete(LPSAFEARRAY* custom)
{
	HRESULT hr = E_FAIL;
	if (_loader && _loader->IsLoaded())
	{
		hr = _loader->OuterAggregator()->Addin()->OnStartupComplete(custom);
	}
	return hr;
}

STDMETHODIMP ShimProxy::OnBeginShutdown(LPSAFEARRAY* custom)
{
	HRESULT hr = E_FAIL;
	if (_loader && _loader->IsLoaded())
	{
		hr = _loader->OuterAggregator()->Addin()->OnBeginShutdown(custom);
	}
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
			ManagedRibbonExtensibility* result = new (std::nothrow) ManagedRibbonExtensibility(ribbon);
			if (result)
			{
				*ppv = static_cast<IRibbonExtensibility*>(result);
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
			ManagedCustomTaskPaneConsumer* result = new (std::nothrow) ManagedCustomTaskPaneConsumer(consumer);
			if (result)
			{
				*ppv = static_cast<ICustomTaskPaneConsumer*>(result);
				hr = S_OK;
			}
			else
			{
				consumer->Release();
				hr = E_OUTOFMEMORY;
			}
		}
	}
	else if (_loader && _loader->IsLoaded())
	{
		// This may cause problems when reload managed addin
		IUnknown* inner = _loader->OuterAggregator()->Addin()->InnerUnkown();
		hr = inner->QueryInterface(riid, ppv);
	}
	else
	{
		hr = E_NOINTERFACE;
	}

	if (NULL != *ppv)
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
