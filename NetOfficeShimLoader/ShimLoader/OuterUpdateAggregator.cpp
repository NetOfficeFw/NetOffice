#include "stdafx.h"
#include "OuterUpdateAggregator.h"


/***************************************************************************
* Ctor Dtor
***************************************************************************/

OuterUpdateAggregator::OuterUpdateAggregator(IShimProxy* parent)
{
	_refCounter = 0;
	_parent = parent;
	_components++;
}

OuterUpdateAggregator::~OuterUpdateAggregator()
{
	_components--;
}


/***************************************************************************
* IOuteUpdateAggregator Implementation
***************************************************************************/

STDMETHODIMP OuterUpdateAggregator::IsAvailable(BOOL* available)
{
	HRESULT hr = S_OK;
	bool result = ENABLE_OUTER_UPDATE_AGGREGATOR && !ENABLE_BLIND_AGGREGATION;
	*available = result;
	return hr;
}

STDMETHODIMP OuterUpdateAggregator::Reload()
{
	HRESULT hr = E_FAIL;
	//MessageBox(GetDesktopWindow(), L"Reload", L"OuterUpdateAggregator", 0);
	if(_parent->IsCLRLoaded())
		IfFailGo(_parent->UnloadCLR());
	IfFailGo(_parent->ReloadCLR());

Error:
	return hr;
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

	if (IID_IUnknown == riid)
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
	{
		hr = E_NOINTERFACE;
	}

	if (NULL != *ppv)
	{
		reinterpret_cast<IUnknown*>(*ppv)->AddRef();
	}

	return hr;
}

STDMETHODIMP_(ULONG) OuterUpdateAggregator::AddRef(void)
{
	_refCounter++;
	return _refCounter;
}

STDMETHODIMP_(ULONG) OuterUpdateAggregator::Release(void)
{
	_refCounter--;
	return _refCounter;
}
