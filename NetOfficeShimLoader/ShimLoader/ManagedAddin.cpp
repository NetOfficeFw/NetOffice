#include "stdafx.h"
#include "ManagedAddin.h"


/***************************************************************************
* Ctor, Dtor
***************************************************************************/

ManagedAddin::ManagedAddin(IUnknown* innerUnkown)
{
	_refCounter = 0;
	_innerUnkown = innerUnkown;
	_components++;
}

ManagedAddin::~ManagedAddin()
{
	if (_innerUnkown)
	{
		_innerUnkown->Release();
		_innerUnkown = nullptr;

	}
	_components--;
}


/***************************************************************************
* IDTExtensibility2 Implementation
***************************************************************************/

STDMETHODIMP ManagedAddin::OnConnection(IDispatch* application, ext_ConnectMode connectMode, IDispatch* addInInst, LPSAFEARRAY* custom)
{
	IDTExtensibility2* extensibility = nullptr;
	HRESULT hr = _innerUnkown->QueryInterface(__uuidof(IDTExtensibility2), (LPVOID*)&extensibility);
	if (hr == S_OK)
	{
		hr = extensibility->OnConnection(application, connectMode, addInInst, custom);
		extensibility->Release();
	}
	return hr;
}

STDMETHODIMP ManagedAddin::OnDisconnection(ext_DisconnectMode removeMode, LPSAFEARRAY* custom)
{
	IDTExtensibility2* extensibility = nullptr;
	HRESULT hr = _innerUnkown->QueryInterface(__uuidof(IDTExtensibility2), (LPVOID*)&extensibility);
	if (hr == S_OK)
	{
		hr = extensibility->OnDisconnection(removeMode, custom);
		extensibility->Release();
	}

	return hr;
}

STDMETHODIMP ManagedAddin::OnAddInsUpdate(LPSAFEARRAY* custom)
{
	IDTExtensibility2* extensibility = nullptr;
	HRESULT hr = _innerUnkown->QueryInterface(__uuidof(IDTExtensibility2), (LPVOID*)&extensibility);
	if (hr == S_OK)
	{
		hr = extensibility->OnAddInsUpdate(custom);
		extensibility->Release();
	}
	return hr;
}

STDMETHODIMP ManagedAddin::OnStartupComplete(LPSAFEARRAY* custom)
{
	IDTExtensibility2* extensibility = nullptr;
	HRESULT hr = _innerUnkown->QueryInterface(__uuidof(IDTExtensibility2), (LPVOID*)&extensibility);
	if (hr == S_OK)
	{
		hr = extensibility->OnStartupComplete(custom);
		extensibility->Release();
	}
	return hr;
}

STDMETHODIMP ManagedAddin::OnBeginShutdown(LPSAFEARRAY* custom)
{
	IDTExtensibility2* extensibility = nullptr;
	HRESULT hr = _innerUnkown->QueryInterface(__uuidof(IDTExtensibility2), (LPVOID*)&extensibility);
	if (hr == S_OK)
	{
		hr = extensibility->OnBeginShutdown(custom);
		extensibility->Release();
	}
	return hr;
}


/***************************************************************************
* IRibbonExtensibility Implementation
***************************************************************************/

STDMETHODIMP ManagedAddin::GetCustomUI(BSTR RibbonID, BSTR* RibbonXml)
{
	IRibbonExtensibility* ribbon = nullptr;
	HRESULT hr = _innerUnkown->QueryInterface(__uuidof(IRibbonExtensibility), (LPVOID*)&ribbon);
	if (hr == S_OK)
	{
		hr = ribbon->GetCustomUI(RibbonID, RibbonXml);
		ribbon->Release();
	}
	return hr;
}


/***************************************************************************
* ICustomTaskPaneConsumer Implementation
***************************************************************************/

STDMETHODIMP ManagedAddin::CTPFactoryAvailable(ICTPFactory* CTPFactoryInst)
{
	ICustomTaskPaneConsumer* paneConsumer = nullptr;
	HRESULT hr = _innerUnkown->QueryInterface(__uuidof(ICustomTaskPaneConsumer), (LPVOID*)&paneConsumer);
	if (hr == S_OK)
	{
		hr = paneConsumer->CTPFactoryAvailable(CTPFactoryInst);
		paneConsumer->Release();
	}
	return hr;
}


/***************************************************************************
* IDispatch Implementation
***************************************************************************/

STDMETHODIMP ManagedAddin::GetTypeInfoCount(UINT* pctinfo)
{
	IDispatch* dispatch = nullptr;
	HRESULT hr = _innerUnkown->QueryInterface(__uuidof(IDispatch), (LPVOID*)&dispatch);
	if (hr == S_OK)
	{
		hr = dispatch->GetTypeInfoCount(pctinfo);
		dispatch->Release();
	}
	return hr;
}

STDMETHODIMP ManagedAddin::GetTypeInfo(UINT iTInfo, LCID lcid, ITypeInfo** ppTInfo)
{
	IDispatch* dispatch = nullptr;
	HRESULT hr = _innerUnkown->QueryInterface(__uuidof(IDispatch), (LPVOID*)&dispatch);
	if (hr == S_OK)
	{
		hr = dispatch->GetTypeInfo(iTInfo, lcid, ppTInfo);
		dispatch->Release();
	}
	return hr;
}

STDMETHODIMP ManagedAddin::GetIDsOfNames(REFIID riid, LPOLESTR* rgszNames, UINT cNames, LCID lcid, DISPID* rgDispId)
{
	IDispatch* dispatch = nullptr;
	HRESULT hr = _innerUnkown->QueryInterface(__uuidof(IDispatch), (LPVOID*)&dispatch);
	if (hr == S_OK)
	{
		hr = dispatch->GetIDsOfNames(riid, rgszNames, cNames, lcid, rgDispId);
		dispatch->Release();
	}
	return hr;
}

STDMETHODIMP ManagedAddin::Invoke(DISPID dispIdMember, REFIID riid, LCID lcid, WORD wFlags, DISPPARAMS* pDispParams, VARIANT* pVarResult, EXCEPINFO* pExcepInfo, UINT* puArgErr)
{
	IDispatch* dispatch = nullptr;
	HRESULT hr = _innerUnkown->QueryInterface(__uuidof(IDispatch), (LPVOID*)&dispatch);
	if (hr == S_OK)
	{
		hr = dispatch->Invoke(dispIdMember, riid, lcid, wFlags, pDispParams, pVarResult, pExcepInfo, puArgErr);
		dispatch->Release();
	}
	return hr;
}


/***************************************************************************
* IUnknown Implementation
***************************************************************************/

STDMETHODIMP ManagedAddin::QueryInterface(REFIID riid, void** ppv)
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

STDMETHODIMP_(ULONG) ManagedAddin::AddRef(void)
{
	return ++_refCounter;
}

STDMETHODIMP_(ULONG) ManagedAddin::Release(void)
{
	if (_refCounter > 0)
		_refCounter--;
	return _refCounter;
}
