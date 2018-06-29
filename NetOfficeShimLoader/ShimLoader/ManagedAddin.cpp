#include "stdafx.h"
#include "ManagedAddin.h"


/***************************************************************************
* Ctor, Dtor
***************************************************************************/

ManagedAddin::ManagedAddin(IUnknown* innerUnkown)
{
	_refCounter = 0;
	_components++;
	_innerUnkown = innerUnkown;
	 _innerUnkown->QueryInterface(__uuidof(IDTExtensibility2), (LPVOID*)&this->_innerExtensibility);
	 _innerUnkown->QueryInterface(__uuidof(IDispatch), (LPVOID*)&this->_innerDispatch);
	 _innerUnkown->QueryInterface(__uuidof(IRibbonExtensibility), (LPVOID*)&this->_innerRibbonExtensibility);
	 _innerUnkown->QueryInterface(__uuidof(ICustomTaskPaneConsumer), (LPVOID*)&this->_innerTaskPaneConsumer);
}

ManagedAddin::~ManagedAddin()
{
	if (_innerTaskPaneConsumer)
	{
		_innerTaskPaneConsumer->Release();
		_innerTaskPaneConsumer = nullptr;
	}
	if (_innerRibbonExtensibility)
	{
		_innerRibbonExtensibility->Release();
		_innerRibbonExtensibility = nullptr;
	}
	if (_innerDispatch)
	{
		_innerDispatch->Release();
		_innerDispatch = nullptr;
	}
	if (_innerExtensibility)
	{
		_innerExtensibility->Release();
		_innerExtensibility = nullptr;
	}
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
	HRESULT hr = E_FAIL;
	if(_innerExtensibility)
		hr = _innerExtensibility->OnConnection(application, connectMode, addInInst, custom);
	return hr;
}

STDMETHODIMP ManagedAddin::OnDisconnection(ext_DisconnectMode removeMode, LPSAFEARRAY* custom)
{
	HRESULT hr = E_FAIL;
	if (_innerExtensibility)
		hr = _innerExtensibility->OnDisconnection(removeMode, custom);
	return hr;
}

STDMETHODIMP ManagedAddin::OnAddInsUpdate(LPSAFEARRAY* custom)
{
	HRESULT hr = E_FAIL;
	if (_innerExtensibility)
		hr = _innerExtensibility->OnAddInsUpdate(custom);
	return hr;
}

STDMETHODIMP ManagedAddin::OnStartupComplete(LPSAFEARRAY* custom)
{
	HRESULT hr = E_FAIL;
	if (_innerExtensibility)
		hr = _innerExtensibility->OnStartupComplete(custom);
	return hr;
}

STDMETHODIMP ManagedAddin::OnBeginShutdown(LPSAFEARRAY* custom)
{
	HRESULT hr = E_FAIL;
	if (_innerExtensibility)
		hr = _innerExtensibility->OnBeginShutdown(custom);
	return hr;
}


/***************************************************************************
* IRibbonExtensibility Implementation
***************************************************************************/

STDMETHODIMP ManagedAddin::GetCustomUI(BSTR RibbonID, BSTR* RibbonXml)
{
	HRESULT hr = E_FAIL;
	if (_innerRibbonExtensibility)
		hr = _innerRibbonExtensibility->GetCustomUI(RibbonID, RibbonXml);
	return hr;
}


/***************************************************************************
* ICustomTaskPaneConsumer Implementation
***************************************************************************/

STDMETHODIMP ManagedAddin::CTPFactoryAvailable(ICTPFactory* CTPFactoryInst)
{
	HRESULT hr = E_FAIL;
	if (_innerTaskPaneConsumer)
		hr = _innerTaskPaneConsumer->CTPFactoryAvailable(CTPFactoryInst);
	return hr;
}


/***************************************************************************
* IDispatch Implementation
***************************************************************************/

STDMETHODIMP ManagedAddin::GetTypeInfoCount(UINT* pctinfo)
{
	HRESULT hr = E_FAIL;
	if (_innerDispatch)
		hr = _innerDispatch->GetTypeInfoCount(pctinfo);
	return hr;
}

STDMETHODIMP ManagedAddin::GetTypeInfo(UINT iTInfo, LCID lcid, ITypeInfo** ppTInfo)
{
	HRESULT hr = E_FAIL;
	if (_innerDispatch)
		hr = _innerDispatch->GetTypeInfo(iTInfo, lcid, ppTInfo);
	return hr;
}

STDMETHODIMP ManagedAddin::GetIDsOfNames(REFIID riid, LPOLESTR* rgszNames, UINT cNames, LCID lcid, DISPID* rgDispId)
{
	HRESULT hr = E_FAIL;
	if (_innerDispatch)
		hr = _innerDispatch->GetIDsOfNames(riid, rgszNames, cNames, lcid, rgDispId);
	return hr;
}

STDMETHODIMP ManagedAddin::Invoke(DISPID dispIdMember, REFIID riid, LCID lcid, WORD wFlags, DISPPARAMS* pDispParams, VARIANT* pVarResult, EXCEPINFO* pExcepInfo, UINT* puArgErr)
{
	HRESULT hr = E_FAIL;
	if (_innerDispatch)
		hr = _innerDispatch->Invoke(dispIdMember, riid, lcid, wFlags, pDispParams, pVarResult, pExcepInfo, puArgErr);
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
