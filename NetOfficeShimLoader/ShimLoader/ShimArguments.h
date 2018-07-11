#pragma once
#include "StdAfx.h"
#include "metahost.h"
#include "atlbase.h"
#include "strsafe.h"

extern HANDLE		_thread;
extern HINSTANCE	_module;
extern void IncComponents(LPCWSTR type);
extern void DecComponents(LPCWSTR type);

namespace NetOffice_ShimLoader
{
	class ShimArguments
	{

	public:

		ShimArguments();
		virtual ~ShimArguments();

		HRESULT Load();
		HRESULT Unload();

	protected:

		HRESULT LoadManagedAddin(MSXML::IXMLDOMDocument2Ptr docPtr);
		HRESULT LoadManagedAggregator(MSXML::IXMLDOMDocument2Ptr docPtr);
		HRESULT LoadAppDomain(MSXML::IXMLDOMDocument2Ptr docPtr);
		HRESULT AppendPath(LPWSTR pszPath, LPCWSTR pszMore);

	};
}
