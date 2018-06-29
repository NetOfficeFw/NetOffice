#include "StdAfx.h"
#include "ShimProxy.h"


/***************************************************************************
* Ctor Dtor
***************************************************************************/

ShimProxy::ShimProxy()
{
	_refCounter = 0;
	_loader = NULL;
	_components++;
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
/*	long lResult = 5;
	WCHAR szBuffer[10];
	wsprintf(szBuffer, L"%ld", lResult);
	MessageBox(GetDesktopWindow(), szBuffer, L"lResult", 0);*/

	if (NULL == application)
		return E_POINTER;

	HRESULT hr = S_OK;

	_loader = new (std::nothrow) ClrHost();
	IfNullGo(_loader);

	IfFailGo(_loader->Load());

	hr = _loader->Aggregator()->Addin()->OnConnection(application, connectMode, addInInst, custom);

	return hr;

Error:

	if (_loader)
	{
		delete _loader;
		_loader = nullptr;
	}

	return hr;
}

STDMETHODIMP ShimProxy::OnDisconnection(ext_DisconnectMode removeMode, LPSAFEARRAY* custom)
{
	HRESULT hr = S_OK;

	if (_loader)
	{
		hr = _loader->Aggregator()->Addin()->OnDisconnection(removeMode, custom);

		_loader->Unload();
		delete _loader;
		_loader = NULL;
	}

	return hr;
}

STDMETHODIMP ShimProxy::OnAddInsUpdate(LPSAFEARRAY* custom)
{
	HRESULT hr = S_OK;

	if (_loader)
	{
		hr = _loader->Aggregator()->Addin()->OnAddInsUpdate(custom);
	}

	return hr;
}

STDMETHODIMP ShimProxy::OnStartupComplete(LPSAFEARRAY* custom)
{
	HRESULT hr = S_OK;

	if (_loader)
	{
		hr = _loader->Aggregator()->Addin()->OnStartupComplete(custom);
	}

	return hr;
}

STDMETHODIMP ShimProxy::OnBeginShutdown(LPSAFEARRAY* custom)
{
	HRESULT hr = S_OK;

	if (_loader)
	{
		hr = _loader->Aggregator()->Addin()->OnBeginShutdown(custom);
	}

	return hr;
}


/***************************************************************************
* IRibbonExtensibility Implementation
***************************************************************************/

STDMETHODIMP ShimProxy::GetCustomUI(BSTR RibbonID, BSTR* RibbonXml)
{
	HRESULT hr = S_OK;

	if (_loader)
	{
		hr = _loader->Aggregator()->Addin()->GetCustomUI(RibbonID, RibbonXml);
	}

	return hr;
}


/***************************************************************************
* ICustomTaskPaneConsumer Implementation
***************************************************************************/

STDMETHODIMP ShimProxy::CTPFactoryAvailable(ICTPFactory* CTPFactoryInst)
{
	HRESULT hr = S_OK;

	if (_loader)
	{
		hr = _loader->Aggregator()->Addin()->CTPFactoryAvailable(CTPFactoryInst);
	}

	return hr;
}


/***************************************************************************
* IDispatch Implementation
***************************************************************************/

STDMETHODIMP ShimProxy::GetTypeInfoCount(UINT* pctinfo)
{
	HRESULT hr = S_OK;

	if (_loader)
	{
		hr = _loader->Aggregator()->Addin()->GetTypeInfoCount(pctinfo);
	}

	return hr;
}

STDMETHODIMP ShimProxy::GetTypeInfo(UINT iTInfo, LCID lcid, ITypeInfo** ppTInfo)
{
	HRESULT hr = S_OK;

	if (_loader)
	{
		hr = _loader->Aggregator()->Addin()->GetTypeInfo(iTInfo, lcid, ppTInfo);
	}

	return hr;
}

STDMETHODIMP ShimProxy::GetIDsOfNames(REFIID riid, LPOLESTR* rgszNames, UINT cNames, LCID lcid, DISPID* rgDispId)
{
	HRESULT hr = S_OK;

	if (_loader)
	{
		hr = _loader->Aggregator()->Addin()->GetIDsOfNames(riid, rgszNames, cNames, lcid, rgDispId);
	}

	return hr;
}

STDMETHODIMP ShimProxy::Invoke(DISPID dispIdMember, REFIID riid, LCID lcid, WORD wFlags, DISPPARAMS* pDispParams, VARIANT* pVarResult, EXCEPINFO* pExcepInfo, UINT* puArgErr)
{
	HRESULT hr = S_OK;

	if (_loader)
	{
		hr = _loader->Aggregator()->Addin()->Invoke(dispIdMember, riid, lcid, wFlags, pDispParams, pVarResult, pExcepInfo, puArgErr);
	}

	return hr;
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

	if ((IID_IDTExtensibility2 == riid) || (IID_IUnknown == riid) || (IID_IDispatch == riid))
	{
		*ppv = static_cast<IDTExtensibility2*>(this);
		hr = S_OK;
	}
	else if ((__uuidof(IRibbonExtensibility) == riid))
	{
		*ppv = static_cast<IRibbonExtensibility*>(this);
		hr = S_OK;
	}
	else if ((__uuidof(ICustomTaskPaneConsumer) == riid))
	{
		*ppv = static_cast<ICustomTaskPaneConsumer*>(this);
		hr = S_OK;
	}
	else
		hr = E_NOINTERFACE;

	if (NULL != *ppv)
	{
		reinterpret_cast<IUnknown*>(*ppv)->AddRef();
	}

	return hr;
}

STDMETHODIMP_(ULONG) ShimProxy::AddRef(void)
{
	return ++_refCounter;
}

STDMETHODIMP_(ULONG) ShimProxy::Release(void)
{
	if (0 != --_refCounter)
		return _refCounter;
	--_components;
	delete this;
	return 0;
}