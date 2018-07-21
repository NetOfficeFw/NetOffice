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

extern HANDLE		_thread;
extern HINSTANCE	_module;
extern void IncComponents(LPCWSTR type);
extern void DecComponents(LPCWSTR type);

namespace NetOffice_ShimLoader
{
	class ShimArguments
	{

	public:

		// Ctor, Dtor
		ShimArguments();
		virtual ~ShimArguments();

		// ShimArguments Methods
		BOOL IsLoaded();
		HRESULT Load();
		HRESULT Unload();
		HRESULT ReadState();
		HRESULT Read();

	protected:

		HRESULT AppendPath(LPWSTR pszPath, LPCWSTR pszMore);

	private:

		HRESULT ReadShimRegister(MSXML::IXMLDOMNodePtr document);
		HRESULT ReadShimSettings(MSXML::IXMLDOMNodePtr document);
		HRESULT ReadShimDefaults(MSXML::IXMLDOMNodePtr document);
		HRESULT ReadManagedAddinAggregator(MSXML::IXMLDOMNodePtr document);
		HRESULT ReadManagedUpdateAggregator(MSXML::IXMLDOMNodePtr document);

		MSXML::IXMLDOMDocument2Ptr LoadFile(WCHAR* fileName);

		HRESULT ComInitialize();
		HRESULT ComUninitialize();
		BOOL ToBool(_bstr_t value);

		WCHAR							_documentPath[MAX_PATH +1];
		MSXML::IXMLDOMDocument2Ptr		_document;
		bool							_coInitialized;
		HRESULT							_readState;
	};
}
