#include "stdafx.h"
#include "ManagedRibbonExtensibility.h"
#include "Vars.h"

namespace NetOffice_ShimLoader
{
	/***************************************************************************
	* Ctor, Dtor
	***************************************************************************/

	ManagedRibbonExtensibility::ManagedRibbonExtensibility(IShimProxy* parent, IRibbonExtensibility* innerExtensibility)
	{
		_refCounter = 0;
		_parent = parent;
		SetInnerPointer(innerExtensibility);
		IncComponents(L"ManagedRibbonExtensibility");
	}

	ManagedRibbonExtensibility::~ManagedRibbonExtensibility()
	{
		if (_innerExtensibility)
		{
			_innerExtensibility->Release();
			_innerExtensibility = nullptr;
		}
		DecComponents(L"ManagedRibbonExtensibility");
	}


	/***************************************************************************
	* ManagedRibbonExtensibility Methods
	***************************************************************************/

	STDMETHODIMP ManagedRibbonExtensibility::SetInnerPointer(IRibbonExtensibility* innerExtensibility)
	{
		HRESULT hr = E_FAIL;

		if (innerExtensibility)
		{
			_innerExtensibility = innerExtensibility;
			hr = S_OK;
		}
		else
		{
			hr = E_POINTER;
		}
		return hr;
	}


	/***************************************************************************
	* IRibbonExtensibility Implementation
	***************************************************************************/

	STDMETHODIMP ManagedRibbonExtensibility::GetCustomUI(BSTR RibbonID, BSTR* RibbonXml)
	{
		HRESULT hr = E_FAIL;
		if (_parent && !_parent->IsReloadThreadInProgress() && _innerExtensibility)
		{
			hr = _innerExtensibility->GetCustomUI(RibbonID, RibbonXml);
		}
		return hr;
	}


	/***************************************************************************
	* IDispatch Implementation
	***************************************************************************/

	STDMETHODIMP ManagedRibbonExtensibility::GetTypeInfoCount(UINT* pctinfo)
	{
		HRESULT hr = E_FAIL;
		IDispatch* dispatch = nullptr;

		if (_parent && !_parent->IsReloadThreadInProgress() && _innerExtensibility)
		{
			hr = _innerExtensibility->QueryInterface(IID_IDispatch, (LPVOID*)&dispatch);
			if (SUCCEEDED(hr))
			{
				hr = dispatch->GetTypeInfoCount(pctinfo);
				dispatch->Release();
			}
		}

		return hr;
	}

	STDMETHODIMP ManagedRibbonExtensibility::GetTypeInfo(UINT iTInfo, LCID lcid, ITypeInfo** ppTInfo)
	{
		HRESULT hr = E_FAIL;
		IDispatch* dispatch = nullptr;

		if (_parent && !_parent->IsReloadThreadInProgress() && _innerExtensibility)
		{
			hr = _innerExtensibility->QueryInterface(IID_IDispatch, (LPVOID*)&dispatch);
			if (SUCCEEDED(hr))
			{
				hr = dispatch->GetTypeInfo(iTInfo, lcid, ppTInfo);
				dispatch->Release();
			}
		}

		return hr;
	}

	STDMETHODIMP ManagedRibbonExtensibility::GetIDsOfNames(REFIID riid, LPOLESTR* rgszNames, UINT cNames, LCID lcid, DISPID* rgDispId)
	{
		HRESULT hr = E_FAIL;
		IDispatch* dispatch = nullptr;

		if (_parent && !_parent->IsReloadThreadInProgress() && _innerExtensibility)
		{
			hr = _innerExtensibility->QueryInterface(IID_IDispatch, (LPVOID*)&dispatch);
			if (SUCCEEDED(hr))
			{
				hr = dispatch->GetIDsOfNames(riid, rgszNames, cNames, lcid, rgDispId);
				dispatch->Release();
			}
		}

		return hr;
	}

	STDMETHODIMP ManagedRibbonExtensibility::Invoke(DISPID dispIdMember, REFIID riid, LCID lcid, WORD wFlags, DISPPARAMS* pDispParams, VARIANT* pVarResult, EXCEPINFO* pExcepInfo, UINT* puArgErr)
	{
		HRESULT hr = E_FAIL;
		IDispatch* dispatch = nullptr;

		if (_parent && !_parent->IsReloadThreadInProgress() && _innerExtensibility)
		{
			hr = _innerExtensibility->QueryInterface(IID_IDispatch, (LPVOID*)&dispatch);
			if (SUCCEEDED(hr))
			{
				hr = dispatch->Invoke(dispIdMember, riid, lcid, wFlags, pDispParams, pVarResult, pExcepInfo, puArgErr);
				dispatch->Release();
			}
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
		bool isBlind = false;
		bool available = _parent && !_parent->IsReloadThreadInProgress();

		if (IID_IUnknown == riid && available)
		{
			*ppv = static_cast<IUnknown*>(this);
			hr = S_OK;
		}
		else if (IID_IDispatch == riid && available)
		{
			*ppv = static_cast<IDispatch*>(this);
			hr = S_OK;
		}
		else if ((IID_IRibbonExtensibility == riid && available))
		{
			*ppv = static_cast<IRibbonExtensibility*>(this);
			hr = S_OK;
		}
		else if (!ENABLE_OUTER_UPDATE_AGGREGATOR && ENABLE_BLIND_AGGREGATION && available)
		{
			hr = _innerExtensibility->QueryInterface(riid, ppv);
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

	STDMETHODIMP_(ULONG) ManagedRibbonExtensibility::AddRef(void)
	{
		_refCounter++;
		return _refCounter;
	}

	STDMETHODIMP_(ULONG) ManagedRibbonExtensibility::Release(void)
	{
		_refCounter--;
		return _refCounter;
	}
}
