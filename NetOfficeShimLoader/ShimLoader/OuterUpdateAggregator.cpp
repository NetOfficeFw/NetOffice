#include "stdafx.h"
#include "OuterUpdateAggregator.h"

namespace NetOffice_ShimLoader
{

	/***************************************************************************
	* Ctor, Dtor
	***************************************************************************/

	OuterUpdateAggregator::OuterUpdateAggregator()
	{
		_refCounter = 0;
		IncComponents(L"OuterUpdateAggregator");
	}

	OuterUpdateAggregator::~OuterUpdateAggregator()
	{
		if (_innerHandler)
		{
			delete _innerHandler;
			_innerHandler = nullptr;
		}
		DecComponents(L"OuterUpdateAggregator");
	}


	/***************************************************************************
	* OuterUpdateAggregator Methods
	***************************************************************************/

	ManagedUpdateHandler* OuterUpdateAggregator::ManagedUpdater()
	{
		return _innerHandler;
	}


	/***************************************************************************
	* IOuterComAggregator Implementation
	***************************************************************************/

	HRESULT __stdcall OuterUpdateAggregator::SetInnerHandler(IUnknown* innerHandler)
	{
		if (innerHandler == NULL)
			return E_POINTER;
		if (_innerHandler != NULL)
			return E_UNEXPECTED;

		_innerHandler = new ManagedUpdateHandler(innerHandler);

		return S_OK;
	}


	/***************************************************************************
	* IUnknown Implementation
	***************************************************************************/

	STDMETHODIMP OuterUpdateAggregator::QueryInterface(REFIID riid, void** ppv)
	{
		if (NULL == ppv)
			return E_POINTER;
		*ppv = NULL;

		HRESULT hr = E_FAIL;

		if ((IID_IUnknown == riid))
		{
			*ppv = static_cast<IUnknown*>(this);
			hr = S_OK;
		}
		else if ((IID_IOuterUpdateAggregator == riid))
		{
			*ppv = static_cast<IOuterUpdateAggregator*>(this);
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

	STDMETHODIMP_(ULONG) OuterUpdateAggregator::AddRef(void)
	{
		return ++_refCounter;
	}

	STDMETHODIMP_(ULONG) OuterUpdateAggregator::Release(void)
	{
		if (_refCounter > 0)
			_refCounter--;
		return _refCounter;
	}
}
