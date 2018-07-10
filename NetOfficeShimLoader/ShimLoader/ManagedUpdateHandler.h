#pragma once
#include "stdafx.h"
#include "Aggregators.h"

extern HANDLE		_thread;
extern HINSTANCE	_module;
extern void IncComponents(LPCWSTR type);
extern void DecComponents(LPCWSTR type);

using namespace NetOffice_Tools_Isolation;

namespace NetOffice_ShimLoader
{
	class ManagedUpdateHandler
	{

	public:

		// Ctor, Dtor
		ManagedUpdateHandler(IUnknown* innerHandler);
		virtual ~ManagedUpdateHandler();

		// ManagedAddin Methods
		STDMETHODIMP SetApplication(IDispatch* application);
		STDMETHODIMP SetCustomData(BSTR custom);
		BOOL STDMETHODCALLTYPE CanExecute();
		STDMETHODIMP Execute();
		STDMETHODIMP Close();

		IManagedInnerUpdateHandler* STDMETHODCALLTYPE InnerHandler();

	private:

		IManagedInnerUpdateHandler*		_innerHandler;
		ULONG							_refCounter;

	};
}
