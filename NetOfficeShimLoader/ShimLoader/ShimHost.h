#pragma once
#include "stdafx.h"
#include "Aggregators.h"
#include "IShimProxy.hpp"
//#include "Vars.hpp"

extern HANDLE		_thread;
extern HINSTANCE	_module;
extern void IncComponents(LPCWSTR type);
extern void DecComponents(LPCWSTR type);

using namespace NetOffice_Tools_Isolation;

namespace NetOffice_ShimLoader
{
	class ShimHost : public IShimHost
	{

	public:

		// Ctor, Dtor
		ShimHost(IShimProxy* parent);
		virtual ~ShimHost();

		// ShimHost methods
		BSTR STDMETHODCALLTYPE CustomData();
		STDMETHODIMP SetCustomData(BSTR custom);

		// IShimHost Implementation
		STDMETHODIMP IsAvailable(BOOL* available);
		STDMETHODIMP Reload();
		STDMETHODIMP Update(BSTR custom);

		// IUnknown Implementation
		STDMETHODIMP         QueryInterface(REFIID riid, void ** ppv);
		STDMETHODIMP_(ULONG) AddRef(void);
		STDMETHODIMP_(ULONG) Release(void);

	private:

		ULONG								_refCounter;
		IShimProxy*							_parent;
		_bstr_t*							_customData;
	};

}
