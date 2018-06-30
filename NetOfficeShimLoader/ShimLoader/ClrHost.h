#pragma once
#include "StdAfx.h"
#include "metahost.h"
#include "atlbase.h"
#include "strsafe.h"
#include "ClrHost.h"
#include "Vars.hpp"
#include "Aggregators.h"
#include "Extensibility2.h"
#include "OuterComAggregator.h"

extern HINSTANCE _module;
extern ULONG _components;
extern ULONG _locks;

class ClrHost : public IOuterUpdateAggregator
{

public:

	// Ctor, Dtor
	ClrHost();
	~ClrHost();

	// ClrLoader Methods
	bool IsLoaded();
	OuterComAggregator* OuterAggregator();
	HRESULT Load();
	HRESULT Unload();

	// IOuterComAggregator Implementation
	STDMETHODIMP Reload();

	// IUnknown Implementation
	STDMETHODIMP         QueryInterface(REFIID riid, void ** ppv);
	STDMETHODIMP_(ULONG) AddRef(void);
	STDMETHODIMP_(ULONG) Release(void);

protected:

	HRESULT AppendPath(LPWSTR pszPath, LPCWSTR pszMore);

private:

	ICorRuntimeHost*			_runtimeHost;
	mscorlib::_AppDomain*		_appDomain;
	OuterComAggregator*			_outerAggregator;
	bool						_isLoaded;
	ULONG						_refCounter;

};
