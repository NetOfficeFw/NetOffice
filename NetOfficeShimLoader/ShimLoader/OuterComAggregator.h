#pragma once
#include "Aggregators.h"
#include "ManagedAddin.h"
#include "ManagedRibbonExtensibility.h"
#include "CustomRegisterValue.h"

extern HANDLE		_thread;
extern HINSTANCE	_module;
extern void IncComponents(LPCWSTR type);
extern void DecComponents(LPCWSTR type);

using namespace NetOffice_Tools_Isolation;

namespace NetOffice_ShimLoader
{
	class OuterComAggregator : public IOuterComAggregator
	{

	public:

		// Ctor, Dtor
		OuterComAggregator();
		virtual ~OuterComAggregator();

		// OuterComAggregator Methods
		ManagedAddin* Addin();

		// IOuterComAggregator Implementation
		HRESULT __stdcall SetInnerAddin(IUnknown *innerAddin);

		// IUnknown Implementation
		STDMETHODIMP         QueryInterface(REFIID riid, void ** ppv);
		STDMETHODIMP_(ULONG) AddRef(void);
		STDMETHODIMP_(ULONG) Release(void);

	private:

		ManagedAddin*		_innerAddin;
		ULONG				_refCounter;

	};
}
