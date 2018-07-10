#pragma once
#include "stdafx.h"
#include "Aggregators.h"
#include "ManagedUpdateHandler.h"

using namespace NetOffice_Tools_Isolation;

extern HANDLE		_thread;
extern HINSTANCE	_module;
extern void IncComponents(LPCWSTR type);
extern void DecComponents(LPCWSTR type);

namespace NetOffice_ShimLoader
{
	class OuterUpdateAggregator : public IOuterUpdateAggregator
	{

	public:

		// Ctor, Dtor
		OuterUpdateAggregator();
		virtual ~OuterUpdateAggregator();

		// OuterUpdateAggregator
		ManagedUpdateHandler* ManagedUpdater();

		// IOuterUpdateAggregator Implementation
		STDMETHODIMP SetInnerHandler(IUnknown* innerHandler);

		// IUnknown Implementation
		STDMETHODIMP         QueryInterface(REFIID riid, void ** ppv);
		STDMETHODIMP_(ULONG) AddRef(void);
		STDMETHODIMP_(ULONG) Release(void);

	private:

		ULONG						_refCounter;
		ManagedUpdateHandler*		_innerHandler;
	};
}
