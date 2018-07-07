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

extern HANDLE		_thread;
extern HINSTANCE	_module;
extern ULONG		_components;
extern ULONG		_locks;

namespace NetOffice_ShimLoader
{
	class ClrHost
	{

	public:

		// Ctor, Dtor
		ClrHost(IOuterUpdateAggregator* updateAggregator);
		~ClrHost();

		// ClrLoader Methods
		BOOL IsLoaded();
		OuterComAggregator* OuterAggregator();
		HRESULT Load();
		HRESULT Unload();
		HRESULT LastLoadError();

	protected:

		HRESULT AppendPath(LPWSTR pszPath, LPCWSTR pszMore);

	private:

		ICorRuntimeHost * _runtimeHost;
		mscorlib::_AppDomain*		_appDomain;
		OuterComAggregator*			_outerAggregator;
		IOuterUpdateAggregator*		_outerUpdateAggregator;
		BOOL						_isLoaded;
		IOuterUpdateAggregator*		_updateAggregator;
		HRESULT						_lastLoadError;

	};

}
