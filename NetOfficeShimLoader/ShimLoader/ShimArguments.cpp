#include "stdafx.h"
#include "ShimArguments.h"
#include "string.h"
#include "wchar.h"
#include <fstream>
#include <iostream>
#include <string>
#include <msxml.h>
#include "Vars.hpp"

using namespace std;

namespace NetOffice_ShimLoader
{
	HRESULT GetDllDirectory(TCHAR *szPath, DWORD nPathBufferSize);

	ShimArguments::ShimArguments()
	{
		_document = nullptr;
		_coInitialized = false;
		IncComponents(L"ShimArguments");
	}

	ShimArguments::~ShimArguments()
	{
		Unload();
		DecComponents(L"ShimArguments");
	}

	BOOL ShimArguments::IsLoaded()
	{
		return NULL != _document;
	}

	HRESULT ShimArguments::Load()
	{
		HRESULT hr = E_FAIL;
		bool b = FALSE;

		WCHAR directoryPath[MAX_PATH + 1];
		IfFailGo(GetDllDirectory(directoryPath, ARRAYSIZE(directoryPath)));

		WCHAR moduleFileName[MAX_PATH + 1];
		IfFailGo(GetModuleFileName(_module, moduleFileName, ARRAYSIZE(moduleFileName)));

		WCHAR fullSettingsFilePath[MAX_PATH + 1];
		IfFailGo(AppendPath(fullSettingsFilePath, directoryPath));
		IfFailGo(AppendPath(fullSettingsFilePath, moduleFileName));
		PWSTR target = StrCatBuff(fullSettingsFilePath, L".ShimSettings", ARRAYSIZE(fullSettingsFilePath));

		IfFalseGo(PathFileExists(target));


		hr = ComInitialize();
		hr = _document.CreateInstance(__uuidof(MSXML::DOMDocument60), NULL, CLSCTX_INPROC_SERVER);
		hr = VARIANT_TRUE == _document->load(target) ? S_OK : E_FAIL;

		//if (VARIANT_TRUE == _document->load(target))
		//{
		//	IfFailGo(_document->setProperty("SelectionLanguage", "XPath"));
		//	IfFailGo(LoadManagedAggregator(_document));
		//	IfFailGo(LoadAppDomain(_document));
		//	IfFailGo(LoadManagedAddin(_document));
		//}
		//_document.Release();



		hr = S_OK;

	Error:

		if (_document)
			_document.Release();

		return hr;
	}

	HRESULT ShimArguments::Unload()
	{
		HRESULT hr = S_OK;
		if (_document)
		{
			_document.Release();
			_document = nullptr;
			hr = ComUninitialize();
		}
		else
		{
			hr = E_FAIL;
		}
		return hr;
	}

	HRESULT ShimArguments::ReadRegisterArguments()
	{
		HRESULT hr = E_FAIL;

		if (IsLoaded())
		{
			MSXML::IXMLDOMNodePtr registerMode = _document->selectSingleNode("/ShimLoader/Shim/Register/Mode");


		}
		else
		{
			hr = E_FAIL;
		}

		return hr;
	}

	HRESULT ShimArguments::ComInitialize()
	{
		HRESULT hr = E_FAIL;

		if (!_coInitialized)
		{
			hr = CoInitialize(NULL);
			if (SUCCEEDED(hr))
				_coInitialized = true;
		}

		return hr;
	}

	HRESULT ShimArguments::ComUninitialize()
	{
		HRESULT hr = E_FAIL;

		if (_coInitialized)
		{
			CoUninitialize();
			_coInitialized = false;
			hr = S_OK;
		}
		return hr;
	}

	HRESULT ShimArguments::LoadManagedAddin(MSXML::IXMLDOMDocument2Ptr docPtr)
	{
		HRESULT hr = S_OK;

		MSXML::IXMLDOMNodePtr assemblyName = nullptr;
		MSXML::IXMLDOMNodePtr assemblyFileName = nullptr;
		MSXML::IXMLDOMNodePtr configFileName = nullptr;
		MSXML::IXMLDOMNodePtr className = nullptr;

		assemblyName = docPtr->selectSingleNode("/Root/ManagedAggregator/Target/AssemblyName");
		if (assemblyName)
			Target_AssemblyName = assemblyName->text;
		else
			goto Error;

		assemblyFileName = docPtr->selectSingleNode("/Root/ManagedAggregator/Target/AssemblyFileName");
		if (assemblyFileName)
			Target_AssemblyFileName = assemblyFileName->text;
		else
			goto Error;

		configFileName = docPtr->selectSingleNode("/Root/ManagedAggregator/Target/ConfigFileName");
		if (configFileName)
			Target_ConfigFileName = configFileName->text;
		else
			goto Error;

		className = docPtr->selectSingleNode("/Root/ManagedAggregator/Target/ClassName");
		if (className)
			Target_ConnectClassName = className->text;
		else
			goto Error;

		return hr;

	Error:

		return hr;
	}

	HRESULT ShimArguments::LoadManagedAggregator(MSXML::IXMLDOMDocument2Ptr docPtr)
	{
		HRESULT hr = S_OK;

		MSXML::IXMLDOMNodePtr assemblyName = nullptr;
		MSXML::IXMLDOMNodePtr className = nullptr;

		assemblyName = docPtr->selectSingleNode("/Root/ManagedAggregator/AssemblyName");
		if (assemblyName)
			TargetManagedAggregator_AssemblyName = assemblyName->text;
		else
			goto Error;

		className = docPtr->selectSingleNode("/Root/ManagedAggregator/ClassName");
		if (assemblyName)
			TargetManagedAggregator_ClassName = className->text;
		else
			goto Error;

		return hr;

	Error:

		hr = E_FAIL;
		return hr;
	}

	HRESULT ShimArguments::LoadAppDomain(MSXML::IXMLDOMDocument2Ptr docPtr)
	{
		HRESULT hr = S_OK;

		MSXML::IXMLDOMNodePtr friendlyName = nullptr;
		MSXML::IXMLDOMNodePtr baseFolder = nullptr;

		friendlyName = docPtr->selectSingleNode("/Root/ManagedAggregator/AppDomain/FriendlyName");
		if (friendlyName)
			TargetManagedAggregator_AppDomain_FriendlyName = friendlyName->text;
		else
			goto Error;

		baseFolder = docPtr->selectSingleNode("/Root/ManagedAggregator/AppDomain/BaseFolder");
		if (baseFolder)
			TargetManagedAggregator_AppDomain_BaseFolder = baseFolder->text;
		else
			goto Error;

		return hr;

	Error:

		hr = E_FAIL;
		return hr;
	}


	HRESULT ShimArguments::AppendPath(LPWSTR pszPath, LPCWSTR pszMore)
	{
		HRESULT hr = S_OK;
		if (!PathAppend(pszPath, pszMore))
		{
			hr = E_UNEXPECTED;
		}
		return hr;
	}

	static HRESULT GetDllDirectory(TCHAR *szPath, DWORD nPathBufferSize)
	{
		HMODULE hInstance = _AtlBaseModule.GetModuleInstance();
		if (hInstance == 0)
		{
			return E_FAIL;
		}

		TCHAR szModule[MAX_PATH + 1];
		DWORD dwFLen = ::GetModuleFileName(hInstance, szModule, MAX_PATH);
		if (dwFLen == 0)
		{
			return E_FAIL;
		}

		TCHAR* pszFileName;
		dwFLen = ::GetFullPathName(
			szModule, nPathBufferSize, szPath, &pszFileName);
		if (dwFLen == 0 || dwFLen >= nPathBufferSize)
		{
			return E_FAIL;
		}

		*pszFileName = 0;
		return S_OK;
	}
}
