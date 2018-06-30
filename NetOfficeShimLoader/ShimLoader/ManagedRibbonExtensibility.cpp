#include "stdafx.h"
#include "ManagedRibbonExtensibility.h"


/***************************************************************************
* Ctor, Dtor
***************************************************************************/

ManagedRibbonExtensibility::ManagedRibbonExtensibility(IUnknown* innerUnkown)
{
	_refCounter = 0;
	_innerUnkown = innerUnkown;
	_components++;
}

ManagedRibbonExtensibility::~ManagedRibbonExtensibility()
{
	if (_innerUnkown)
	{
		_innerUnkown->Release();
		_innerUnkown = nullptr;

	}
	_components--;
}


/***************************************************************************
* IRibbonExtensibility Implementation
***************************************************************************/

STDMETHODIMP ManagedRibbonExtensibility::GetCustomUI(BSTR RibbonID, BSTR* RibbonXml)
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
* IDispatch Implementation
***************************************************************************/

STDMETHODIMP ManagedRibbonExtensibility::GetTypeInfoCount(UINT* pctinfo)
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

STDMETHODIMP ManagedRibbonExtensibility::GetTypeInfo(UINT iTInfo, LCID lcid, ITypeInfo** ppTInfo)
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

STDMETHODIMP ManagedRibbonExtensibility::GetIDsOfNames(REFIID riid, LPOLESTR* rgszNames, UINT cNames, LCID lcid, DISPID* rgDispId)
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

STDMETHODIMP ManagedRibbonExtensibility::Invoke(DISPID dispIdMember, REFIID riid, LCID lcid, WORD wFlags, DISPPARAMS* pDispParams, VARIANT* pVarResult, EXCEPINFO* pExcepInfo, UINT* puArgErr)
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

STDMETHODIMP ManagedRibbonExtensibility::QueryInterface(REFIID riid, void** ppv)
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
	else if ((__uuidof(IRibbonExtensibility) == riid))
	{
		*ppv = static_cast<IRibbonExtensibility*>(this);
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

STDMETHODIMP_(ULONG) ManagedRibbonExtensibility::AddRef(void)
{
	_refCounter++;
	return _refCounter;
}

STDMETHODIMP_(ULONG) ManagedRibbonExtensibility::Release(void)
{
	if (_refCounter > 0)
		_refCounter--;
	return _refCounter;
}
