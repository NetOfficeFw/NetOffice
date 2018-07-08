#include "stdafx.h"
#include "ManagedAddin.h"

using namespace NetOffice_Tools_Isolation;

namespace NetOffice_ShimLoader
{
	/***************************************************************************
	* Ctor, Dtor
	***************************************************************************/

	ManagedAddin::ManagedAddin(IUnknown* innerUnkown)
	{
		_refCounter = 0;
		_innerUnkown = innerUnkown;
		IncComponents(L"ManagedAddin");
	}

	ManagedAddin::~ManagedAddin()
	{
		if (_innerUnkown)
		{
			_innerUnkown->Release();
			_innerUnkown = nullptr;

		}
		DecComponents(L"ManagedAddin");
	}


	/***************************************************************************
	* ManagedAddin Methods
	***************************************************************************/

	IUnknown* ManagedAddin::InnerUnkown()
	{
		return _innerUnkown;
	}

	STDMETHODIMP ManagedAddin::ReloadNotification(BSTR custom)
	{
		HRESULT hr = E_FAIL;

		IManagedInnerAddin* inner = nullptr;

		if (_innerUnkown)
		{
			hr = _innerUnkown->QueryInterface(IID_IManagedInnerAddin, (LPVOID*)&inner);
			if (SUCCEEDED(hr))
			{
				hr = inner->ReloadNotification(custom);
				inner->Release();
			}
		}

		return hr;
	}

	/***************************************************************************
	* IDTExtensibility2 Implementation
	***************************************************************************/

	STDMETHODIMP ManagedAddin::OnConnection(IDispatch* application, ext_ConnectMode connectMode, IDispatch* addInInst, LPSAFEARRAY* custom)
	{
		HRESULT hr = E_FAIL;
		IDTExtensibility2* extensibility = nullptr;

		if (_innerUnkown)
		{
			hr = _innerUnkown->QueryInterface(IID_IDTExtensibility2, (LPVOID*)&extensibility);
			if (SUCCEEDED(hr))
			{
				hr = extensibility->OnConnection(application, connectMode, addInInst, custom);
				extensibility->Release();
			}
		}

		return hr;
	}

	STDMETHODIMP ManagedAddin::OnDisconnection(ext_DisconnectMode removeMode, LPSAFEARRAY* custom)
	{
		HRESULT hr = E_FAIL;
		IDTExtensibility2* extensibility = nullptr;

		if (_innerUnkown)
		{
			hr = _innerUnkown->QueryInterface(IID_IDTExtensibility2, (LPVOID*)&extensibility);
			if (SUCCEEDED(hr))
			{
				hr = extensibility->OnDisconnection(removeMode, custom);
				extensibility->Release();
			}
		}

		return hr;
	}

	STDMETHODIMP ManagedAddin::OnAddInsUpdate(LPSAFEARRAY* custom)
	{
		HRESULT hr = E_FAIL;
		IDTExtensibility2* extensibility = nullptr;

		if (_innerUnkown)
		{
			hr = _innerUnkown->QueryInterface(IID_IDTExtensibility2, (LPVOID*)&extensibility);
			if (SUCCEEDED(hr))
			{
				hr = extensibility->OnAddInsUpdate(custom);
				extensibility->Release();
			}
		}

		return hr;
	}

	STDMETHODIMP ManagedAddin::OnStartupComplete(LPSAFEARRAY* custom)
	{
		HRESULT hr = E_FAIL;
		IDTExtensibility2* extensibility = nullptr;

		if (_innerUnkown)
		{
			hr = _innerUnkown->QueryInterface(IID_IDTExtensibility2, (LPVOID*)&extensibility);
			if (SUCCEEDED(hr))
			{
				hr = extensibility->OnStartupComplete(custom);
				extensibility->Release();
			}
		}

		return hr;
	}

	STDMETHODIMP ManagedAddin::OnBeginShutdown(LPSAFEARRAY* custom)
	{
		HRESULT hr = E_FAIL;
		IDTExtensibility2* extensibility = nullptr;

		if (_innerUnkown)
		{
			hr = _innerUnkown->QueryInterface(IID_IDTExtensibility2, (LPVOID*)&extensibility);
			if (SUCCEEDED(hr))
			{
				hr = extensibility->OnBeginShutdown(custom);
				extensibility->Release();
			}
		}

		return hr;
	}


	/***************************************************************************
	* IDispatch Implementation
	***************************************************************************/

	STDMETHODIMP ManagedAddin::GetTypeInfoCount(UINT* pctinfo)
	{
		IDispatch* dispatch = nullptr;
		HRESULT hr = _innerUnkown->QueryInterface(IID_IDispatch, (LPVOID*)&dispatch);
		if (SUCCEEDED(hr))
		{
			hr = dispatch->GetTypeInfoCount(pctinfo);
			dispatch->Release();
		}
		return hr;
	}

	STDMETHODIMP ManagedAddin::GetTypeInfo(UINT iTInfo, LCID lcid, ITypeInfo** ppTInfo)
	{
		IDispatch* dispatch = nullptr;
		HRESULT hr = _innerUnkown->QueryInterface(IID_IDispatch, (LPVOID*)&dispatch);
		if (SUCCEEDED(hr))
		{
			hr = dispatch->GetTypeInfo(iTInfo, lcid, ppTInfo);
			dispatch->Release();
		}
		return hr;
	}

	STDMETHODIMP ManagedAddin::GetIDsOfNames(REFIID riid, LPOLESTR* rgszNames, UINT cNames, LCID lcid, DISPID* rgDispId)
	{
		IDispatch* dispatch = nullptr;
		HRESULT hr = _innerUnkown->QueryInterface(IID_IDispatch, (LPVOID*)&dispatch);
		if (SUCCEEDED(hr))
		{
			hr = dispatch->GetIDsOfNames(riid, rgszNames, cNames, lcid, rgDispId);
			dispatch->Release();
		}
		return hr;
	}

	STDMETHODIMP ManagedAddin::Invoke(DISPID dispIdMember, REFIID riid, LCID lcid, WORD wFlags, DISPPARAMS* pDispParams, VARIANT* pVarResult, EXCEPINFO* pExcepInfo, UINT* puArgErr)
	{
		IDispatch* dispatch = nullptr;
		HRESULT hr = _innerUnkown->QueryInterface(IID_IDispatch, (LPVOID*)&dispatch);
		if (SUCCEEDED(hr))
		{
			hr = dispatch->Invoke(dispIdMember, riid, lcid, wFlags, pDispParams, pVarResult, pExcepInfo, puArgErr);
			dispatch->Release();
		}
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

	STDMETHODIMP_(ULONG) ManagedAddin::AddRef(void)
	{
		_refCounter++;
		return _refCounter;
	}

	STDMETHODIMP_(ULONG) ManagedAddin::Release(void)
	{
		if (_refCounter > 0)
			_refCounter--;
		return _refCounter;
	}
}
