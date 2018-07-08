#include "stdafx.h"
#include "ManagedUpdateHandler.h"

namespace NetOffice_ShimLoader
{
	/***************************************************************************
	* Ctor, Dtor
	***************************************************************************/

	ManagedUpdateHandler::ManagedUpdateHandler(IUnknown* innerHandler)
	{
		innerHandler->QueryInterface(IID_IManagedInnerUpdateHandler, (LPVOID*)&_innerHandler);
		IncComponents(L"ManagedUpdateHandler");
	}

	ManagedUpdateHandler::~ManagedUpdateHandler()
	{
		if (_innerHandler)
		{
			_innerHandler->Release();
			_innerHandler = nullptr;
		}
		DecComponents(L"ManagedUpdateHandler");
	}


	/***************************************************************************
	* ManagedUpdateHandler Methods
	***************************************************************************/

	STDMETHODIMP ManagedUpdateHandler::SetApplication(IDispatch* application)
	{
		HRESULT hr = E_FAIL;

		if (_innerHandler)
		{
			hr = _innerHandler->SetApplication(application);
		}
		return hr;
	}

	STDMETHODIMP ManagedUpdateHandler::SetCustomData(BSTR custom)
	{
		HRESULT hr = E_FAIL;

		if (_innerHandler)
		{
			hr = _innerHandler->SetCustomData(custom);
		}
		return hr;
	}

	BOOL STDMETHODCALLTYPE ManagedUpdateHandler::CanExecute()
	{
		BOOL result = FALSE;
		BOOL canExecute = FALSE;

		if (_innerHandler)
		{
			HRESULT hr = _innerHandler->CanExecute(&canExecute);
			if(SUCCEEDED(hr))
				result = canExecute;
		}

		return result;
	}

	STDMETHODIMP ManagedUpdateHandler::Execute()
	{
		HRESULT hr = E_FAIL;

		if (_innerHandler)
		{
			hr = _innerHandler->Execute();
		}
		return hr;
	}

	IManagedInnerUpdateHandler* STDMETHODCALLTYPE ManagedUpdateHandler::InnerHandler()
	{
		return _innerHandler;
	}

	STDMETHODIMP ManagedUpdateHandler::Close()
	{
		HRESULT hr = E_FAIL;

		if (_innerHandler)
		{
			hr = _innerHandler->Close();
		}
		return hr;
	}
}
