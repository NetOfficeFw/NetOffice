#pragma once
#include "StdAfx.h"
#include "metahost.h"
#include "atlbase.h"
#include "strsafe.h"
#include "Aggregators.h"
#include "Extensibility2.h"
#include "OuterComAggregator.h"
//#include "Vars.hpp"

extern HANDLE		_thread;
extern HINSTANCE	_module;
extern void IncComponents(LPCWSTR type);
extern void DecComponents(LPCWSTR type);

namespace NetOffice_ShimLoader
{
	class ClrHost
	{

	public:

		// Ctor, Dtor
		ClrHost(IShimHost* shimHost);
		virtual ~ClrHost();

		// ClrLoader Methods
		BOOL IsLoaded();
		OuterComAggregator* OuterAggregator();
		HRESULT Load();
		HRESULT Unload();
		HRESULT LastLoadError();

	protected:

		HRESULT AppendPath(LPWSTR pszPath, LPCWSTR pszMore);

	private:

		ICorRuntimeHost*			_runtimeHost;
		mscorlib::_AppDomain*		_appDomain;
		OuterComAggregator*			_outerAggregator;
		IShimHost*					_shimHost;
		BOOL						_isLoaded;
		HRESULT						_lastLoadError;

	};
}
