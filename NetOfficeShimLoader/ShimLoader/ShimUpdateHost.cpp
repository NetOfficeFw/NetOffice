#include "stdafx.h"
#include "ShimUpdateHost.h"

namespace NetOffice_ShimLoader
{
	/***************************************************************************
	* Ctor, Dtor
	***************************************************************************/

	ShimUpdateHost::ShimUpdateHost()
	{
		_refCounter = 0;
		_customData = nullptr;
		IncComponents(L"ShimUpdateHost");
	}

	ShimUpdateHost::~ShimUpdateHost()
	{
		if (_customData)
		{
			delete _customData;
			_customData = nullptr;
		}
		DecComponents(L"ShimUpdateHost");
	}


	/***************************************************************************
	* ShimUpdateHost Methods
	***************************************************************************/

	BSTR ShimUpdateHost::CustomData()
	{
		return  NULL != _customData ? _customData->copy(true) : NULL;
	}


	/***************************************************************************
	* IShimUpdateHost Implementation
	***************************************************************************/

	STDMETHODIMP ShimUpdateHost::SetCustomData(BSTR custom)
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

	STDMETHODIMP ShimUpdateHost::Done()
	{
		HRESULT hr = S_OK;

		return hr;
	}


	/***************************************************************************
	* IUnknown Implementation
	***************************************************************************/

	STDMETHODIMP ShimUpdateHost::QueryInterface(REFIID riid, void** ppv)
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
		else if ((IID_IShimUpdateHost == riid))
		{
			*ppv = static_cast<IShimUpdateHost*>(this);
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

	STDMETHODIMP_(ULONG) ShimUpdateHost::AddRef(void)
	{
		_refCounter++;
		return _refCounter;
	}

	STDMETHODIMP_(ULONG) ShimUpdateHost::Release(void)
	{
		_refCounter--;
		return _refCounter;
	}
}
