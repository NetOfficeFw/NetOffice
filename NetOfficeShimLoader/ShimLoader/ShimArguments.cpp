#include "stdafx.h"
#include "ShimArguments.h"
#include "Vars.h"
#include "PathParser.h"

using namespace std;
using namespace NetOffice_ShimLoader_Register;

namespace NetOffice_ShimLoader
{
	// TODO: check if needed
	HRESULT GetDllDirectory(TCHAR *szPath, DWORD nPathBufferSize);

	/***************************************************************************
	* Ctor, Dtor
	***************************************************************************/

	ShimArguments::ShimArguments()
	{
		_document = nullptr;
		_coInitialized = false;
		_readState = E_NOT_SET;
		IncComponents(L"ShimArguments");
	}

	ShimArguments::~ShimArguments()
	{
		Unload();
		ComUninitialize();
		DecComponents(L"ShimArguments");
	}


	/***************************************************************************
	* ShimArguments Methods
	***************************************************************************/

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

	HRESULT ShimArguments::ReadState()
	{
		return _readState;
	}

	HRESULT ShimArguments::Read()
	{
		HRESULT hr = E_FAIL;
		MSXML::IXMLDOMNodePtr document = nullptr;
		if (IsLoaded())
		{
			hr = _document.QueryInterface(__uuidof(IXMLDOMNode), &document);
			if (SUCCEEDED(hr))
			{
				hr = ReadShimRegister(document);
				if(SUCCEEDED(hr))
					hr = ReadShimSettings(document);
				if (SUCCEEDED(hr))
					hr = ReadShimDefaults(document);
				if (SUCCEEDED(hr))
					hr = ReadManagedAddinAggregator(document);
				if (SUCCEEDED(hr))
					hr = ReadManagedUpdateAggregator(document);
			}
		}
		else
		{
			hr = E_ABORT;
		}

		_readState = hr;
		if (document)
		{
			document.Release();
			document = nullptr;
		}
		return hr;
	}

	HRESULT ShimArguments::ReadShimRegister(MSXML::IXMLDOMNodePtr document)
	{
		HRESULT hr = S_OK;
		DllRegisterModeParser parser;
		MSXML::IXMLDOMNodePtr registerShim = nullptr;
		MSXML::IXMLDOMNodePtr registerTarget = nullptr;
		MSXML::IXMLDOMNodePtr registerAddin = nullptr;
		MSXML::IXMLDOMNodePtr registerMode = nullptr;
		MSXML::IXMLDOMNodePtr registerClsId = nullptr;
		MSXML::IXMLDOMNodePtr registerProgId = nullptr;
		MSXML::IXMLDOMNodePtr friendlyName = nullptr;
		MSXML::IXMLDOMNodePtr description = nullptr;
		MSXML::IXMLDOMNodePtr loadBehavior = nullptr;
		MSXML::IXMLDOMNodePtr commandLineSafe = nullptr;
		MSXML::IXMLDOMNodeListPtr addins = nullptr;
		MSXML::IXMLDOMNodeListPtr customRegs = nullptr;

		registerShim = document->selectSingleNode("/ShimLoader/Shim/Register/RegisterShim");
		IfNullGo(registerShim);
		registerTarget = document->selectSingleNode("/ShimLoader/Shim/Register/RegisterTarget");
		IfNullGo(registerTarget);
		registerAddin = document->selectSingleNode("/ShimLoader/Shim/Register/CreateAddinKeys");
		IfNullGo(registerAddin);
		registerMode = document->selectSingleNode("/ShimLoader/Shim/Register/Mode");
		IfNullGo(registerMode);
		registerClsId = document->selectSingleNode("/ShimLoader/Shim/Register/Component/CLSID");
		IfNullGo(registerClsId);
		registerProgId = document->selectSingleNode("/ShimLoader/Shim/Register/Component/ProgId");
		IfNullGo(registerProgId);
		friendlyName = document->selectSingleNode("/ShimLoader/Shim/Register/Addin/FriendlyName");
		IfNullGo(friendlyName);
		description = document->selectSingleNode("/ShimLoader/Shim/Register/Addin/Description");
		IfNullGo(description);
		loadBehavior = document->selectSingleNode("/ShimLoader/Shim/Register/Addin/LoadBehavior");
		IfNullGo(loadBehavior);
		commandLineSafe = document->selectSingleNode("/ShimLoader/Shim/Register/Addin/CommandLineSafe");
		IfNullGo(commandLineSafe);
		addins = document->selectNodes("/ShimLoader/Shim/Register/Addin/Applications/*");
		IfNullGo(addins);

		ShimProxy_Host_Application_Length = addins->length;
		ShimProxy_Host_Application = new LPCWSTR[addins->length];
		for (int i = 0; i < addins->length; i++)
		{
			MSXML::IXMLDOMNode* domNode = nullptr;
			if (SUCCEEDED(addins->get_item(i, &domNode)))
			{
				_bstr_t appName = domNode->GetnodeName();
				ShimProxy_Host_Application[i] = appName.copy(TRUE);
				domNode->Release();
			}
		}

		ENABLE_SELF_REGISTRATION = ToBool(registerShim->text);
		ENABLE_TARGET_REGISTRATION = ToBool(registerTarget->text);
		ENABLE_ADDIN_REGISTRATION = ToBool(registerAddin->text);
		lstrcpyn(ShimProxy_CLSID, registerClsId->text, MAX_PATH + 1);
		lstrcpyn(ShimProxy_ProgID, registerProgId->text, MAX_PATH + 1);
		lstrcpyn(ShimProxy_FriendlyName, friendlyName->text, MAX_PATH + 1);
		lstrcpyn(ShimProxy_Description, description->text, MAX_PATH + 1);

		auto mode = parser.Parse(registerMode->text);
		SELF_REGISTER_MODE = mode;

		customRegs = document->selectNodes("/ShimLoader/Shim/Register/Addin/CustomRegs/CustomReg");
		if (customRegs)
		{
			Custom_Register_Values_Length = customRegs->length;
			Custom_Register_Values = new PCustomRegisterValue[customRegs->length];
			for (int i = 0; i < customRegs->length; i++)
			{
				MSXML::IXMLDOMNode* domNode = nullptr;
				if (SUCCEEDED(customRegs->get_item(i, &domNode)))
				{
					auto nameNode = domNode->selectSingleNode("Name");
					auto typeNode = domNode->selectSingleNode("Type");
					auto valueNode = domNode->selectSingleNode("Value");
					auto parseValueNode = domNode->selectSingleNode("ParseValue");

					if (nameNode && typeNode && valueNode)
					{
						auto theName = nameNode->text.copy(true);
						auto theType = typeNode->text.copy(true);
						auto theValue = valueNode->text.copy(true);
						auto parseTheValue = NULL != parseValueNode ? ToBool(parseValueNode->text) : FALSE;
						Custom_Register_Values[i] = new CustomRegisterValue(theName, theType, theValue, parseTheValue);
					}
					domNode->Release();
				}
			}
		}

		return hr;

	Error:

		hr = E_FAIL;
		return hr;
	}

	HRESULT ShimArguments::ReadShimSettings(MSXML::IXMLDOMNodePtr document)
	{
		HRESULT hr = S_OK;
		MSXML::IXMLDOMNodePtr enabledNode = nullptr;
		MSXML::IXMLDOMNodePtr blindAggEnabledNode = nullptr;
		MSXML::IXMLDOMNodePtr updateEnabledNode = nullptr;
		MSXML::IXMLDOMNodePtr debugMessageBoxNode = nullptr;

		enabledNode = document->selectSingleNode("/ShimLoader/Shim/Settings/Enabled");
		IfNullGo(enabledNode);
		blindAggEnabledNode = document->selectSingleNode("/ShimLoader/Shim/Settings/BlindAggregationEnabled");
		IfNullGo(blindAggEnabledNode);
		updateEnabledNode = document->selectSingleNode("/ShimLoader/Shim/Settings/UpdateEnabled");
		IfNullGo(updateEnabledNode);
		debugMessageBoxNode = document->selectSingleNode("/ShimLoader/Shim/Settings/DebugMsgBoxEnabled");

		ENABLE_SHIM = ToBool(enabledNode->text);
		ENABLE_BLIND_AGGREGATION = ToBool(blindAggEnabledNode->text);
		ENABLE_OUTER_UPDATE_AGGREGATOR = ToBool(updateEnabledNode->text);
		if (debugMessageBoxNode)
			ENABLE_DEBUG_MESSAGE_BOX = ToBool(debugMessageBoxNode->text);

		return hr;

	Error:

		hr = E_FAIL;
		return hr;
	}

	HRESULT ShimArguments::ReadShimDefaults(MSXML::IXMLDOMNodePtr document)
	{
		HRESULT hr = S_OK;
		MSXML::IXMLDOMNodePtr extensibilityDefaultNode = nullptr;
		MSXML::IXMLDOMNodePtr extensibilityFailNode = nullptr;

		extensibilityDefaultNode = document->selectSingleNode("/ShimLoader/Shim/Defaults/ExtensibilityDefaultResult");
		IfNullGo(extensibilityDefaultNode);
		extensibilityFailNode = document->selectSingleNode("/ShimLoader/Shim/Defaults/ExtensibilityFailResult");
		IfNullGo(extensibilityFailNode);

		EXTENSIBILITY_DEFAULT_RESULT = stol(extensibilityDefaultNode->text.copy());
		EXTENSIBILITY_FAIL_RESULT = stol(extensibilityFailNode->text.copy());

		return hr;

	Error:

		hr = E_FAIL;
		return hr;
	}

	HRESULT ShimArguments::ReadManagedAddinAggregator(MSXML::IXMLDOMNodePtr document)
	{
		HRESULT hr = S_OK;
		PathParser parser;
		MSXML::IXMLDOMNodePtr folderPathNode = nullptr;
		MSXML::IXMLDOMNodePtr folderPathSubFolderNode = nullptr;
		MSXML::IXMLDOMNodePtr assemblyNameNode = nullptr;
		MSXML::IXMLDOMNodePtr classNameNode = nullptr;

		MSXML::IXMLDOMNodePtr appDomainFriendlyNameNode = nullptr;
		MSXML::IXMLDOMNodePtr appDomainFolderPathNode = nullptr;
		MSXML::IXMLDOMNodePtr appDomainSubFolderNode = nullptr;

		MSXML::IXMLDOMNodePtr targetAssemblyFileName = nullptr;
		MSXML::IXMLDOMNodePtr targetAssemblyName = nullptr;
		MSXML::IXMLDOMNodePtr targetConfigFileName = nullptr;
		MSXML::IXMLDOMNodePtr targetClassName = nullptr;

		folderPathNode = document->selectSingleNode("/ShimLoader/ManagedAddinAggregator/Folder/Path");
		IfNullGo(folderPathNode);
		folderPathSubFolderNode = document->selectSingleNode("/ShimLoader/ManagedAddinAggregator/Folder/SubFolder");
		IfNullGo(folderPathSubFolderNode);
		assemblyNameNode = document->selectSingleNode("/ShimLoader/ManagedAddinAggregator/AssemblyName");
		IfNullGo(assemblyNameNode);
		classNameNode = document->selectSingleNode("/ShimLoader/ManagedAddinAggregator/ClassName");
		IfNullGo(classNameNode);
		appDomainFriendlyNameNode = document->selectSingleNode("/ShimLoader/ManagedAddinAggregator/AppDomain/FriendlyName");
		IfNullGo(appDomainFriendlyNameNode);
		appDomainFolderPathNode = document->selectSingleNode("/ShimLoader/ManagedAddinAggregator/AppDomain/Folder/Path");
		IfNullGo(appDomainFolderPathNode);
		appDomainSubFolderNode = document->selectSingleNode("/ShimLoader/ManagedAddinAggregator/AppDomain/Folder/SubFolder");
		IfNullGo(appDomainSubFolderNode);
		targetAssemblyFileName = document->selectSingleNode("/ShimLoader/ManagedAddinAggregator/Target/AssemblyFileName");
		IfNullGo(targetAssemblyFileName);
		targetAssemblyName = document->selectSingleNode("/ShimLoader/ManagedAddinAggregator/Target/AssemblyName");
		IfNullGo(targetAssemblyName);
		targetConfigFileName = document->selectSingleNode("/ShimLoader/ManagedAddinAggregator/Target/ConfigFileName");
		IfNullGo(targetConfigFileName);
		targetClassName = document->selectSingleNode("/ShimLoader/ManagedAddinAggregator/Target/ClassName");
		IfNullGo(targetClassName);

		IfFailGo(parser.ParseEx(folderPathNode->text, folderPathSubFolderNode->text, TargetManagedAggregator_Folder, MAX_PATH + 1));
		IfFailGo(parser.ParseEx(appDomainFolderPathNode->text, appDomainSubFolderNode->text, TargetManagedAggregator_AppDomain_BaseFolder, MAX_PATH + 1));

		lstrcpyn(TargetManagedAggregator_AssemblyName, assemblyNameNode->text, MAX_PATH + 1);
		lstrcpyn(TargetManagedAggregator_ClassName, classNameNode->text, MAX_PATH + 1);
		lstrcpyn(TargetManagedAggregator_AppDomain_FriendlyName, appDomainFriendlyNameNode->text, MAX_PATH + 1);
		lstrcpyn(Target_AssemblyName, targetAssemblyName->text, MAX_PATH + 1);
		lstrcpyn(Target_AssemblyFileName, targetAssemblyFileName->text, MAX_PATH + 1);
		lstrcpyn(Target_ConfigFileName, targetConfigFileName->text, MAX_PATH + 1);
		lstrcpyn(Target_ConnectClassName, targetClassName->text, MAX_PATH + 1);

		return hr;

	Error:

		hr = E_FAIL;
		return hr;
	}

	HRESULT ShimArguments::ReadManagedUpdateAggregator(MSXML::IXMLDOMNodePtr document)
	{
		HRESULT hr = S_OK;
		PathParser parser;

		MSXML::IXMLDOMNodePtr folderPathNode = nullptr;
		MSXML::IXMLDOMNodePtr folderPathSubFolderNode = nullptr;
		MSXML::IXMLDOMNodePtr assemblyNameNode = nullptr;
		MSXML::IXMLDOMNodePtr classNameNode = nullptr;

		MSXML::IXMLDOMNodePtr appDomainFriendlyNameNode = nullptr;
		MSXML::IXMLDOMNodePtr appDomainFolderPathNode = nullptr;
		MSXML::IXMLDOMNodePtr appDomainSubFolderNode = nullptr;

		MSXML::IXMLDOMNodePtr targetAssemblyFileName = nullptr;
		MSXML::IXMLDOMNodePtr targetAssemblyName = nullptr;
		MSXML::IXMLDOMNodePtr targetConfigFileName = nullptr;
		MSXML::IXMLDOMNodePtr targetClassName = nullptr;

		folderPathNode = document->selectSingleNode("/ShimLoader/ManagedUpdateAggregator/Folder/Path");
		IfNullGo(folderPathNode);
		folderPathSubFolderNode = document->selectSingleNode("/ShimLoader/ManagedUpdateAggregator/Folder/SubFolder");
		IfNullGo(folderPathSubFolderNode);
		assemblyNameNode = document->selectSingleNode("/ShimLoader/ManagedUpdateAggregator/AssemblyName");
		IfNullGo(assemblyNameNode);
		classNameNode = document->selectSingleNode("/ShimLoader/ManagedUpdateAggregator/ClassName");
		IfNullGo(classNameNode);

		appDomainFriendlyNameNode = document->selectSingleNode("/ShimLoader/ManagedUpdateAggregator/AppDomain/FriendlyName");
		IfNullGo(appDomainFriendlyNameNode);
		appDomainFolderPathNode = document->selectSingleNode("/ShimLoader/ManagedUpdateAggregator/AppDomain/Folder/Path");
		IfNullGo(appDomainFolderPathNode);
		appDomainSubFolderNode = document->selectSingleNode("/ShimLoader/ManagedUpdateAggregator/AppDomain/Folder/SubFolder");
		IfNullGo(appDomainSubFolderNode);

		targetAssemblyFileName = document->selectSingleNode("/ShimLoader/ManagedUpdateAggregator/Target/AssemblyFileName");
		IfNullGo(targetAssemblyFileName);
		targetAssemblyName = document->selectSingleNode("/ShimLoader/ManagedUpdateAggregator/Target/AssemblyName");
		IfNullGo(targetAssemblyName);
		targetConfigFileName = document->selectSingleNode("/ShimLoader/ManagedUpdateAggregator/Target/ConfigFileName");
		IfNullGo(targetConfigFileName);
		targetClassName = document->selectSingleNode("/ShimLoader/ManagedUpdateAggregator/Target/ClassName");
		IfNullGo(targetClassName);

		IfFailGo(parser.ParseEx(folderPathNode->text, folderPathSubFolderNode->text, UpdateManagedAggregator_Folder, MAX_PATH + 1));
		IfFailGo(parser.ParseEx(appDomainFolderPathNode->text, appDomainSubFolderNode->text, UpdateManagedAggregator_AppDomain_BaseFolder, MAX_PATH + 1));

		lstrcpyn(UpdateManagedAggregator_AssemblyName, assemblyNameNode->text, MAX_PATH + 1);
		lstrcpyn(UpdateManagedAggregator_ClassName, classNameNode->text, MAX_PATH + 1);
		lstrcpyn(UpdateManagedAggregator_AppDomain_FriendlyName, appDomainFriendlyNameNode->text, MAX_PATH + 1);
		lstrcpyn(Update_AssemblyName, targetAssemblyName->text, MAX_PATH + 1);
		lstrcpyn(Update_AssemblyFileName, targetAssemblyFileName->text, MAX_PATH + 1);
		lstrcpyn(Update_ConfigFileName, targetConfigFileName->text, MAX_PATH + 1);
		lstrcpyn(Update_ConnectClassName, targetClassName->text, MAX_PATH + 1);

		return hr;

	Error:

		hr = E_FAIL;
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
