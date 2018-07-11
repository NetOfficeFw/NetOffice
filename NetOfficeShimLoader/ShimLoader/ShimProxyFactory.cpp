#include "stdafx.h"
#include "ShimProxyFactory.h"

namespace NetOffice_ShimLoader
{
	/***************************************************************************
	* Ctor, Dtor
	***************************************************************************/

	ShimProxyFactory::ShimProxyFactory()
	{
		_refCount = 0;
		IncComponents(L"ShimProxyFactory");
	}

	ShimProxyFactory::~ShimProxyFactory()
	{
		DecComponents(L"ShimProxyFactory");
	}

	/***************************************************************************
	* IClassFactory Implementation
	***************************************************************************/

	STDMETHODIMP ShimProxyFactory::CreateInstance(LPUNKNOWN punk, REFIID riid, void** ppv)
	{
		if (NULL == ppv)
			return E_POINTER;

		*ppv = NULL;

		if (NULL != punk)
			return CLASS_E_NOAGGREGATION;

		if (riid != IID_IDTExtensibility2)
			return E_NOINTERFACE;

		NetOffice_ShimLoader::ShimProxy* pObj = new (std::nothrow) NetOffice_ShimLoader::ShimProxy();
		if (NULL == pObj)
			return E_OUTOFMEMORY;

		HRESULT hr = pObj->QueryInterface(riid, ppv);
		if (!SUCCEEDED(hr))
			delete pObj;

		return hr;
	}

	STDMETHODIMP ShimProxyFactory::LockServer(BOOL fLock)
	{
		if (fLock)
			++_refCount;
		else
			--_refCount;

		return S_OK;
	}


	/***************************************************************************
	* IUnknown Implementation
	***************************************************************************/

	STDMETHODIMP ShimProxyFactory::QueryInterface(REFIID riid, void** ppv)
	{
		if (NULL == ppv)
			return E_POINTER;

		*ppv = NULL;

		HRESULT hr = S_OK;

		if ((IID_IUnknown == riid) || (IID_IClassFactory == riid))
			*ppv = static_cast<IClassFactory*>(this);
		else
			hr = E_NOINTERFACE;

		if (NULL != *ppv)
			reinterpret_cast<IUnknown*>(*ppv)->AddRef();

		return hr;
	}

	STDMETHODIMP_(ULONG) ShimProxyFactory::AddRef(void)
	{
		_refCount++;
		return _refCount;
	}

	STDMETHODIMP_(ULONG) ShimProxyFactory::Release(void)
	{
		_refCount--;
		if (0 == _refCount)
			delete this;
		return 0;
	}
}
