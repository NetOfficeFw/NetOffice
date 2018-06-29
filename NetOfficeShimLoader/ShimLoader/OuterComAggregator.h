#pragma once
#include "Aggregators.h"
#include "ManagedAddin.h"

extern HINSTANCE _module;
extern ULONG _components;
extern ULONG _locks;

class OuterComAggregator : public IOuterComAggregator
{

public:

	// Ctor
	OuterComAggregator();
	~OuterComAggregator();

	// OuterComAggregator Methods
	ManagedAddin* Addin();

	// IOuterComAggregator Implementation
	HRESULT __stdcall SetInnerAddin(IUnknown *innerAddin);

	// IUnknown Implementation
	STDMETHODIMP         QueryInterface(REFIID riid, void ** ppv);
	STDMETHODIMP_(ULONG) AddRef(void);
	STDMETHODIMP_(ULONG) Release(void);

private:

	ManagedAddin* _innerAddin;
	ULONG _refCounter;
};

