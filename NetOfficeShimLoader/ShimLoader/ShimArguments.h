#pragma once
#include "StdAfx.h"
#include "metahost.h"
#include "atlbase.h"
#include "strsafe.h"
#include "string.h"
#include "wchar.h"
#include <fstream>
#include <iostream>
#include <string>
#include <msxml.h>
#include "DllRegisterModeParser.h"
//#include "Vars.hpp"

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

		BOOL IsLoaded();
		HRESULT Load();
		HRESULT Unload();

		HRESULT ReadRegisterArguments();

	protected:

		HRESULT LoadManagedAddin(MSXML::IXMLDOMDocument2Ptr docPtr);
		HRESULT LoadManagedAggregator(MSXML::IXMLDOMDocument2Ptr docPtr);
		HRESULT LoadAppDomain(MSXML::IXMLDOMDocument2Ptr docPtr);
		HRESULT AppendPath(LPWSTR pszPath, LPCWSTR pszMore);

	private:

		HRESULT ComInitialize();
		HRESULT ComUninitialize();

		BOOL ToBool(_bstr_t value);

		MSXML::IXMLDOMDocument2Ptr	_document;
		bool						_coInitialized;
	};
}
