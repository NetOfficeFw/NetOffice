#include "stdafx.h"
#include "ShimHost.h"
#include "Vars.h"

namespace NetOffice_ShimLoader
{
	/***************************************************************************
	* Ctor Dtor
	***************************************************************************/

	ShimHost::ShimHost(IShimProxy* parent)
	{
		_refCounter = 0;
		_parent = parent;
		_customData = nullptr;
		IncComponents(L"ShimHost");
	}

	ShimHost::~ShimHost()
	{
		if (_customData)
		{
			delete _customData;
			_customData = nullptr;
		}
		DecComponents(L"ShimHost");
	}


	/***************************************************************************
	* ShimHost Methods
	***************************************************************************/

	BSTR STDMETHODCALLTYPE ShimHost::CustomData()
	{
		return _customData->copy(true);
	}

	STDMETHODIMP ShimHost::SetCustomData(BSTR custom)
	{
		HRESULT hr = S_OK;
		if (_customData)
		{
			delete _customData;
			_customData = nullptr;
		}
		_customData = new _bstr_t(custom, true);

		return hr;
	}

	/***************************************************************************
	* IShimHost Implementation
	***************************************************************************/

	STDMETHODIMP ShimHost::IsAvailable(BOOL* available)
	{
		HRESULT hr = S_OK;
		bool result = ENABLE_OUTER_UPDATE_AGGREGATOR && !ENABLE_BLIND_AGGREGATION;
		*available = result;
		return hr;
	}

	STDMETHODIMP ShimHost::Reload()
	{
		HRESULT hr = E_FAIL;

		IfFailGo(_parent->ReloadCLR(TRUE));

		return hr;

	Error:
		return hr;
	}

	STDMETHODIMP ShimHost::Update(BSTR custom)
	{
		HRESULT hr = S_OK;

		SetCustomData(custom);
		IfFailGo(_parent->Update(TRUE));

		return hr;

	Error:
		return hr;
	}

	/***************************************************************************
	* IUnknown Implementation
	***************************************************************************/

	STDMETHODIMP ShimHost::QueryInterface(REFIID riid, void** ppv)
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
		else if ((IID_IShimHost == riid))
		{
			*ppv = static_cast<IShimHost*>(this);
			hr = S_OK;
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

	STDMETHODIMP_(ULONG) ShimHost::AddRef(void)
	{
		_refCounter++;
		return _refCounter;
	}

	STDMETHODIMP_(ULONG) ShimHost::Release(void)
	{
		_refCounter--;
		return _refCounter;
	}
}
