#pragma once
#include "StdAfx.h"
#include "clrhost.h"
#include "metahost.h"
#include "atlbase.h"
#include "strsafe.h"
#include "Vars.hpp"
#include "Aggregators.h"
#include "Extensibility2.h"
#include "OuterComAggregator.h"

extern HINSTANCE _module;
extern ULONG _components;
extern ULONG _locks;

class ClrHost
{

public:

	// Ctor, Dtor
	ClrHost();
	~ClrHost();

	// ClrLoader Methods
	OuterComAggregator* Aggregator();
	HRESULT Load();
	HRESULT Unload();

protected:

	HRESULT AppendPath(LPWSTR pszPath, LPCWSTR pszMore);

private:

	ICorRuntimeHost*		_runtimeHost;
	mscorlib::_AppDomain*	_appDomain;
	OuterComAggregator*		_aggregator;
	ULONG					_refCounter;

};
