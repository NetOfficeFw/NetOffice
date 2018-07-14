#include "stdafx.h"
#include "ManagedCustomTaskPaneConsumer.h"
#include "Vars.h"

namespace NetOffice_ShimLoader
{
	/***************************************************************************
	* Ctor, Dtor
	***************************************************************************/

	ManagedCustomTaskPaneConsumer::ManagedCustomTaskPaneConsumer(IShimProxy* parent, ICustomTaskPaneConsumer* innerConsumer)
	{
		_refCounter = 0;
		_parent = parent;
		_ctpFactoryInst = nullptr;
		SetInnerPointer(innerConsumer);
		IncComponents(L"ManagedCustomTaskPaneConsumer");
	}

	ManagedCustomTaskPaneConsumer::~ManagedCustomTaskPaneConsumer()
	{
		if (_innerConsumer)
		{
			_innerConsumer->Release();
			_innerConsumer = nullptr;
		}
		if (_ctpFactoryInst)
		{
			_ctpFactoryInst->Release();
			_ctpFactoryInst = nullptr;
		}
		DecComponents(L"ManagedCustomTaskPaneConsumer");
	}


	/***************************************************************************
	* ManagedCustomTaskPaneConsumer Methods
	***************************************************************************/

	STDMETHODIMP ManagedCustomTaskPaneConsumer::SetInnerPointer(ICustomTaskPaneConsumer* innerConsumer)
	{
		HRESULT hr = E_FAIL;

		if (innerConsumer)
		{
			_innerConsumer = innerConsumer;
			hr = S_OK;
		}
		else
		{
			hr = E_POINTER;
		}
		return hr;
	}

	ICTPFactory* ManagedCustomTaskPaneConsumer::InnerCtpFactory()
	{
		return _ctpFactoryInst;
	}

	/***************************************************************************
	* ICustomTaskPaneConsumer Implementation
	***************************************************************************/

	STDMETHODIMP ManagedCustomTaskPaneConsumer::CTPFactoryAvailable(ICTPFactory* CTPFactoryInst)
	{
		HRESULT hr = E_FAIL;
		ICustomTaskPaneConsumer* paneConsumer = nullptr;
		if (_ctpFactoryInst)
		{
			_ctpFactoryInst->Release();
		}
		_ctpFactoryInst = CTPFactoryInst;
		if (_ctpFactoryInst)
		{
			_ctpFactoryInst->AddRef();
		}

		if (_innerConsumer)
		{
			hr = _innerConsumer->QueryInterface(IID_ICustomTaskPaneConsumer, (LPVOID*)&paneConsumer);
			if (SUCCEEDED(hr))
			{
				hr = paneConsumer->CTPFactoryAvailable(CTPFactoryInst);
				paneConsumer->Release();
			}
		}

		return hr;
	}


	/***************************************************************************
	* IDispatch Implementation
	***************************************************************************/

	STDMETHODIMP ManagedCustomTaskPaneConsumer::GetTypeInfoCount(UINT* pctinfo)
	{
		HRESULT hr = E_FAIL;
		IDispatch* dispatch = nullptr;

		if (_parent && !_parent->IsReloadThreadInProgress() && _innerConsumer)
		{
			hr = _innerConsumer->QueryInterface(IID_IDispatch, (LPVOID*)&dispatch);
			if (SUCCEEDED(hr))
			{
				hr = dispatch->GetTypeInfoCount(pctinfo);
				dispatch->Release();
			}
		}

		return hr;
	}

	STDMETHODIMP ManagedCustomTaskPaneConsumer::GetTypeInfo(UINT iTInfo, LCID lcid, ITypeInfo** ppTInfo)
	{
		HRESULT hr = E_FAIL;
		IDispatch* dispatch = nullptr;

		if (_parent && !_parent->IsReloadThreadInProgress() && _innerConsumer)
		{
			hr = _innerConsumer->QueryInterface(IID_IDispatch, (LPVOID*)&dispatch);
			if (SUCCEEDED(hr))
			{
				hr = dispatch->GetTypeInfo(iTInfo, lcid, ppTInfo);
				dispatch->Release();
			}
		}

		return hr;
	}

	STDMETHODIMP ManagedCustomTaskPaneConsumer::GetIDsOfNames(REFIID riid, LPOLESTR* rgszNames, UINT cNames, LCID lcid, DISPID* rgDispId)
	{
		HRESULT hr = E_FAIL;
		IDispatch* dispatch = nullptr;

		if (_parent && !_parent->IsReloadThreadInProgress() && _innerConsumer)
		{
			hr = _innerConsumer->QueryInterface(IID_IDispatch, (LPVOID*)&dispatch);
			if (SUCCEEDED(hr))
			{
				hr = dispatch->GetIDsOfNames(riid, rgszNames, cNames, lcid, rgDispId);
				dispatch->Release();
			}
		}

		return hr;
	}

	STDMETHODIMP ManagedCustomTaskPaneConsumer::Invoke(DISPID dispIdMember, REFIID riid, LCID lcid, WORD wFlags, DISPPARAMS* pDispParams, VARIANT* pVarResult, EXCEPINFO* pExcepInfo, UINT* puArgErr)
	{
		HRESULT hr = E_FAIL;
		IDispatch* dispatch = nullptr;

		if (_parent && !_parent->IsReloadThreadInProgress() && _innerConsumer)
		{
			hr = _innerConsumer->QueryInterface(IID_IDispatch, (LPVOID*)&dispatch);
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

	STDMETHODIMP ManagedCustomTaskPaneConsumer::QueryInterface(REFIID riid, void** ppv)
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
		else if (IID_ICustomTaskPaneConsumer == riid && available)
		{
			*ppv = static_cast<ICustomTaskPaneConsumer*>(this);
			hr = S_OK;
		}
		else if (!ENABLE_OUTER_UPDATE_AGGREGATOR && ENABLE_BLIND_AGGREGATION && available)
		{
			hr = _innerConsumer->QueryInterface(riid, ppv);
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

	STDMETHODIMP_(ULONG) ManagedCustomTaskPaneConsumer::AddRef(void)
	{
		_refCounter++;
		return _refCounter;
	}

	STDMETHODIMP_(ULONG) ManagedCustomTaskPaneConsumer::Release(void)
	{
		_refCounter--;
		return _refCounter;
	}
}
