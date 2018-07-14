#include "stdafx.h"
#include "ShimArguments.h"
#include "Vars.h"

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

		IfFailGo(ComInitialize());
		IfFailGo(_document.CreateInstance(__uuidof(MSXML::DOMDocument60), NULL, CLSCTX_INPROC_SERVER));
		IfFailGo(VARIANT_TRUE == _document->load(target) ? S_OK : E_FAIL);
		IfFailGo(_document->setProperty("SelectionLanguage", "XPath"));

		return hr;

	Error:

		if (_document)
		{
			_document.Release();
			_document = nullptr;
		}

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
		MSXML::IXMLDOMNodePtr document = nullptr;
		MSXML::IXMLDOMNodePtr registerShim = nullptr;
		MSXML::IXMLDOMNodePtr registerTarget = nullptr;
		MSXML::IXMLDOMNodePtr registerMode = nullptr;
		MSXML::IXMLDOMNodePtr registerClsId = nullptr;
		MSXML::IXMLDOMNodePtr registerProgId = nullptr;

		if (IsLoaded())
		{
			hr = _document.QueryInterface(__uuidof(IXMLDOMNode), &document);
			if (SUCCEEDED(hr))
			{
				registerShim = document->selectSingleNode("/ShimLoader/Shim/Register/RegisterShim");
				IfNullGo(registerShim);
				registerTarget = document->selectSingleNode("/ShimLoader/Shim/Register/RegisterTarget");
				IfNullGo(registerTarget);
				registerMode = document->selectSingleNode("/ShimLoader/Shim/Register/Mode");
				IfNullGo(registerMode);
				registerClsId = document->selectSingleNode("/ShimLoader/Shim/Register/Component/CLSID");
				IfNullGo(registerClsId);
				registerProgId = document->selectSingleNode("/ShimLoader/Shim/Register/Component/ProgId");
				IfNullGo(registerProgId);

				ENABLE_SELF_REGISTRATION = ToBool(registerShim->text);
				ENABLE_TARGET_REGISTRATION = ToBool(registerTarget->text);
				ShimProxy_CLSID = registerClsId->text;
				ShimProxy_ProgID = registerProgId->text;
				DllRegisterModeParser parser;
				auto mode = parser.Parse(registerMode->text);
				SELF_REGISTER_MODE = mode;
			}
		}

		if (document)
		{
			document.Release();
			document = nullptr;
		}
		return hr;

	Error:

		if (document)
		{
			document.Release();
			document = nullptr;
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
		HRESULT hr = S_OK;

		if (_coInitialized)
		{
			CoUninitialize();
			_coInitialized = false;
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

		/*assemblyName = docPtr->selectSingleNode("/Root/ManagedAggregator/Target/AssemblyName");
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
			goto Error;*/

		return hr;

	//Error:

	//	return hr;
	}

	HRESULT ShimArguments::LoadManagedAggregator(MSXML::IXMLDOMDocument2Ptr docPtr)
	{
		HRESULT hr = S_OK;

		/*MSXML::IXMLDOMNodePtr assemblyName = nullptr;
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
			goto Error;*/

		return hr;

	//Error:

	//	hr = E_FAIL;
	//	return hr;
	}

	HRESULT ShimArguments::LoadAppDomain(MSXML::IXMLDOMDocument2Ptr docPtr)
	{
		HRESULT hr = S_OK;

		/*MSXML::IXMLDOMNodePtr friendlyName = nullptr;
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
			goto Error;*/

		return hr;

	//Error:

	//	hr = E_FAIL;
	//	return hr;
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

	BOOL ShimArguments::ToBool(_bstr_t value)
	{
		if (0 == wcscmp(value, L"true") ||
			0 == wcscmp(value, L"TRUE") ||
			0 == wcscmp(value, L"True") ||
			0 == wcscmp(value, L"1"))
			return TRUE;
		else
			return FALSE;
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
