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
		IncComponents(L"ShimArguments");
	}

	ShimArguments::~ShimArguments()
	{
		DecComponents(L"ShimArguments");
	}

	HRESULT ShimArguments::Load()
	{
		HRESULT hr = E_FAIL;
		bool initialized = false;
		bool b = FALSE;
		MSXML::IXMLDOMDocument2Ptr docPtr;

		WCHAR directoryPath[MAX_PATH + 1];
		IfFailGo(GetDllDirectory(directoryPath, ARRAYSIZE(directoryPath)));

		WCHAR moduleFileName[MAX_PATH + 1];
		IfFailGo(GetModuleFileName(_module, moduleFileName, ARRAYSIZE(moduleFileName)));

		WCHAR fullSettingsFilePath[MAX_PATH + 1];
		IfFailGo(AppendPath(fullSettingsFilePath, directoryPath));
		IfFailGo(AppendPath(fullSettingsFilePath, moduleFileName));
		PWSTR target = StrCatBuff(fullSettingsFilePath, L".ShimSettings", ARRAYSIZE(fullSettingsFilePath));

		IfFalseGo(PathFileExists(target));


		hr = CoInitialize(NULL);
		initialized = true;
		hr = docPtr.CreateInstance(__uuidof(MSXML::DOMDocument60), NULL, CLSCTX_INPROC_SERVER);
		if (VARIANT_TRUE == docPtr->load(target))
		{
			IfFailGo(docPtr->setProperty("SelectionLanguage", "XPath"));
			IfFailGo(LoadManagedAggregator(docPtr));
			IfFailGo(LoadAppDomain(docPtr));
			IfFailGo(LoadManagedAddin(docPtr));
		}
		docPtr.Release();
		CoUninitialize();


		hr = S_OK;

	Error:
		if (docPtr)
			docPtr.Release();
		if(initialized)
			CoUninitialize();
		return hr;
	}

	HRESULT ShimArguments::LoadManagedAddin(MSXML::IXMLDOMDocument2Ptr docPtr)
	{
		HRESULT hr = S_OK;

		return hr;

	Error:

		hr = E_FAIL;
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
			AppDomain_FriendlyName = friendlyName->text;
		else
			goto Error;

		baseFolder = docPtr->selectSingleNode("/Root/ManagedAggregator/AppDomain/BaseFolder");
		if (baseFolder)
			AppDomain_BaseFolder = baseFolder->text;
		else
			goto Error;

		return hr;

	Error:

		hr = E_FAIL;
		return hr;
	}

	HRESULT ShimArguments::Unload()
	{
		HRESULT hr = S_OK;
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
