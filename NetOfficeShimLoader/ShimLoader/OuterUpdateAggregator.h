#pragma once
#include "stdafx.h"
#include "Aggregators.h"
#include "Vars.hpp"
#include "IShimProxy.hpp"

extern HANDLE		_thread;
extern HINSTANCE	_module;
extern ULONG		_components;
extern ULONG		_locks;

using namespace NetOffice_Tools_Isolation;

namespace NetOffice_ShimLoader
{
	class OuterUpdateAggregator : public IOuterUpdateAggregator
	{

	public:

		// Ctor, Dtor
		OuterUpdateAggregator(IShimProxy* parent);
		~OuterUpdateAggregator();

		// IOuterUpdateAggregator Implementation
		STDMETHODIMP IsAvailable(BOOL* available);
		STDMETHODIMP Reload();

		// IUnknown Implementation
		STDMETHODIMP         QueryInterface(REFIID riid, void ** ppv);
		STDMETHODIMP_(ULONG) AddRef(void);
		STDMETHODIMP_(ULONG) Release(void);

	private:

		ULONG								_refCounter;
		IShimProxy*							_parent;
	};

}
