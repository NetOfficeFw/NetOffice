#pragma once
#include "stdafx.h"
#include "Aggregators.h"
#include "IShimProxy.hpp"

extern HANDLE		_thread;
extern HINSTANCE	_module;
extern void IncComponents(LPCWSTR type);
extern void DecComponents(LPCWSTR type);

using namespace NetOffice_Tools_Isolation;

namespace NetOffice_ShimLoader
{
	class ShimUpdateHost : public IShimUpdateHost
	{

	public:

		// Ctor, Dtor
		ShimUpdateHost();
		virtual ~ShimUpdateHost();

		// ShimUpdateHost Methods
		BSTR CustomData();

		// IShimUpdateHost Implementation
		STDMETHODIMP SetCustomData(BSTR custom);
		STDMETHODIMP Done();

		// IUnknown Implementation
		STDMETHODIMP         QueryInterface(REFIID riid, void ** ppv);
		STDMETHODIMP_(ULONG) AddRef(void);
		STDMETHODIMP_(ULONG) Release(void);

	private:

		ULONG								_refCounter;
		_bstr_t*							_customData;
	};
}
