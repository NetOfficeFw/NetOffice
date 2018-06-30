#include "stdafx.h"
#include "ManagedCustomTaskPaneConsumer.h"


/***************************************************************************
* Ctor, Dtor
***************************************************************************/

ManagedCustomTaskPaneConsumer::ManagedCustomTaskPaneConsumer(ICustomTaskPaneConsumer* innerConsumer)
{
	_refCounter = 0;
	_innerConsumer = innerConsumer;
	_components++;
}


ManagedCustomTaskPaneConsumer::~ManagedCustomTaskPaneConsumer()
{
	if (_innerConsumer)
	{
		_innerConsumer->Release();
		_innerConsumer = nullptr;
	}
	_components--;
}


/***************************************************************************
* ICustomTaskPaneConsumer Implementation
***************************************************************************/

STDMETHODIMP ManagedCustomTaskPaneConsumer::CTPFactoryAvailable(ICTPFactory* CTPFactoryInst)
{
	ICustomTaskPaneConsumer* paneConsumer = nullptr;
	HRESULT hr = _innerConsumer->QueryInterface(__uuidof(ICustomTaskPaneConsumer), (LPVOID*)&paneConsumer);
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

STDMETHODIMP ManagedCustomTaskPaneConsumer::GetTypeInfoCount(UINT* pctinfo)
{
	IDispatch* dispatch = nullptr;
	HRESULT hr = _innerConsumer->QueryInterface(__uuidof(IDispatch), (LPVOID*)&dispatch);
	if (hr == S_OK)
	{
		hr = dispatch->GetTypeInfoCount(pctinfo);
		dispatch->Release();
	}
	return hr;
}

STDMETHODIMP ManagedCustomTaskPaneConsumer::GetTypeInfo(UINT iTInfo, LCID lcid, ITypeInfo** ppTInfo)
{
	IDispatch* dispatch = nullptr;
	HRESULT hr = _innerConsumer->QueryInterface(__uuidof(IDispatch), (LPVOID*)&dispatch);
	if (hr == S_OK)
	{
		hr = dispatch->GetTypeInfo(iTInfo, lcid, ppTInfo);
		dispatch->Release();
	}
	return hr;
}

STDMETHODIMP ManagedCustomTaskPaneConsumer::GetIDsOfNames(REFIID riid, LPOLESTR* rgszNames, UINT cNames, LCID lcid, DISPID* rgDispId)
{
	IDispatch* dispatch = nullptr;
	HRESULT hr = _innerConsumer->QueryInterface(__uuidof(IDispatch), (LPVOID*)&dispatch);
	if (hr == S_OK)
	{
		hr = dispatch->GetIDsOfNames(riid, rgszNames, cNames, lcid, rgDispId);
		dispatch->Release();
	}
	return hr;
}

STDMETHODIMP ManagedCustomTaskPaneConsumer::Invoke(DISPID dispIdMember, REFIID riid, LCID lcid, WORD wFlags, DISPPARAMS* pDispParams, VARIANT* pVarResult, EXCEPINFO* pExcepInfo, UINT* puArgErr)
{
	IDispatch* dispatch = nullptr;
	HRESULT hr = _innerConsumer->QueryInterface(__uuidof(IDispatch), (LPVOID*)&dispatch);
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

STDMETHODIMP ManagedCustomTaskPaneConsumer::QueryInterface(REFIID riid, void** ppv)
{
	if (NULL == ppv)
		return E_POINTER;
	*ppv = NULL;

	HRESULT hr = E_FAIL;

	if (IID_IUnknown == riid)
	{
		*ppv = static_cast<IUnknown*>(this);
		hr = S_OK;
	}
	else if (IID_IDispatch == riid)
	{
		*ppv = static_cast<IDispatch*>(this);
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

STDMETHODIMP_(ULONG) ManagedCustomTaskPaneConsumer::AddRef(void)
{
	_refCounter++;
	return _refCounter;
}

STDMETHODIMP_(ULONG) ManagedCustomTaskPaneConsumer::Release(void)
{
	_refCounter--;
	if (0 == _refCounter)
	{
		delete this;
		return 0;
	}
	else
	{
		return _refCounter;
	}
}
