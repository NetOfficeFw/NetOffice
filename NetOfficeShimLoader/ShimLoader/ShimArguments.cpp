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
		lstrcpyn(_documentPath, L"", MAX_PATH);
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

	MSXML::IXMLDOMDocument2Ptr ShimArguments::LoadFile(WCHAR* fileName)
	{
		NetOffice_ShimLoader_Analytics::WriteLog(L"ShimArguments::LoadFile::Enter");

		HRESULT hr = E_FAIL;
		bool b = FALSE;
		MSXML::IXMLDOMDocument2Ptr document;

		b = PathFileExists(fileName);
		if (!b)
		{
			NetOffice_ShimLoader_Analytics::WriteError(L"ShimArguments::LoadFile::MissingFile", fileName);
		}
		IfFalseGo(b);

		IfFailGo(ComInitialize());
		IfFailGo(document.CreateInstance(__uuidof(MSXML::DOMDocument60), NULL, CLSCTX_INPROC_SERVER));
		IfFailGo(VARIANT_TRUE == document->load(fileName) ? S_OK : E_FAIL);

		IfFailGo(document->setProperty("SelectionLanguage", "XPath"));

		NetOffice_ShimLoader_Analytics::WriteLog(L"ShimArguments::LoadFile::Exit");
		return document;

	Error:

		NetOffice_ShimLoader_Analytics::WriteError(L"ShimArguments::LoadFile::Error", hr);
		if (document)
		{
			document.Release();
			document = nullptr;
		}

		return NULL;
	}

	BOOL ShimArguments::IsLoaded()
	{
		return NULL != _document;
	}

	HRESULT ShimArguments::Load()
	{
		NetOffice_ShimLoader_Analytics::WriteLog(L"ShimArguments::Load::Enter");

		HRESULT hr = E_FAIL;
		MSXML::IXMLDOMNodePtr redirectFolderNode = nullptr;
		MSXML::IXMLDOMNodePtr redirectSubFolderNode = nullptr;
		MSXML::IXMLDOMNodePtr redirectFileNode = nullptr;

		PathParser parser;

		WCHAR directoryPath[MAX_PATH + 1];
		IfFailGo(GetDllDirectory(directoryPath, ARRAYSIZE(directoryPath)));
		lstrcpyn(_documentPath, directoryPath, MAX_PATH);
		WCHAR moduleFileName[MAX_PATH + 1];
		IfFailGo(GetModuleFileName(_module, moduleFileName, ARRAYSIZE(moduleFileName)));

		WCHAR fullSettingsFilePath[MAX_PATH + 1];
		IfFailGo(AppendPath(fullSettingsFilePath, directoryPath));
		IfFailGo(AppendPath(fullSettingsFilePath, moduleFileName));
		PWSTR target = StrCatBuff(fullSettingsFilePath, L".ShimSettings", ARRAYSIZE(fullSettingsFilePath));
		// TODO: delete target then?

		_document = LoadFile(target);

		if (_document)
		{
			for (size_t i = 0; i < 3; i++)
			{
				MSXML::IXMLDOMNodePtr document = nullptr;
				hr = _document.QueryInterface(__uuidof(IXMLDOMNode), &document);
				if (SUCCEEDED(hr))
				{
					redirectFolderNode = document->selectSingleNode("/ShimLoader/Manifest/Folder");
					if (!redirectFolderNode)
						break;
					redirectSubFolderNode = document->selectSingleNode("/ShimLoader/Manifest/SubFolder");
					if (!redirectSubFolderNode)
						break;
					redirectFileNode = document->selectSingleNode("/ShimLoader/Manifest/File");
					if (!redirectFileNode)
						break;

					auto foo1 = redirectFolderNode->text.copy();
					auto foo2 = redirectSubFolderNode->text.copy();
					auto foo3 = redirectFileNode->text.copy();
					parser.ParseEx(foo1, foo2, foo3, _documentPath, MAX_PATH);

					_document = LoadFile(_documentPath);

					document.Release();
					document = nullptr;

					break;
				}
			}
		}
		else
		{
			hr = E_FAIL;
			goto Error;
		}

		NetOffice_ShimLoader_Analytics::WriteLog(L"ShimArguments::Load::Exit");
		return hr;

	Error:

		NetOffice_ShimLoader_Analytics::WriteError(L"ShimArguments::Load::FailExit", hr);

		if (_document)
		{
			_document.Release();
			_document = nullptr;
		}

		return hr;
	}

	HRESULT ShimArguments::Unload()
	{
		NetOffice_ShimLoader_Analytics::WriteLog(L"ShimArguments::Unload::Enter");

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

		NetOffice_ShimLoader_Analytics::WriteLog(L"ShimArguments::Unload::Exit", hr);
		return hr;
	}

	HRESULT ShimArguments::ReadState()
	{
		return _readState;
	}

	HRESULT ShimArguments::Read()
	{
		NetOffice_ShimLoader_Analytics::WriteLog(L"ShimArguments::Read::Enter");

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

		if(SUCCEEDED(hr))
			NetOffice_ShimLoader_Analytics::WriteLog(L"ShimArguments::Read::Exit");
		else
			NetOffice_ShimLoader_Analytics::WriteError(L"ShimArguments::Read::FailExit", hr);

		return hr;
	}

	HRESULT ShimArguments::ReadShimRegister(MSXML::IXMLDOMNodePtr document)
	{
		NetOffice_ShimLoader_Analytics::WriteLog(L"ShimArguments::ReadShimRegister::Enter");

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

		NetOffice_ShimLoader_Analytics::WriteLog(L"ShimArguments::ReadShimRegister::Exit");
		return hr;

	Error:

		hr = E_FAIL;
		NetOffice_ShimLoader_Analytics::WriteLog(L"ShimArguments::ReadShimRegister::FailExit", hr);
		return hr;
	}

	HRESULT ShimArguments::ReadShimSettings(MSXML::IXMLDOMNodePtr document)
	{
		NetOffice_ShimLoader_Analytics::WriteLog(L"ShimArguments::ReadShimSettings::Enter");

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

		NetOffice_ShimLoader_Analytics::WriteLog(L"ShimArguments::ReadShimSettings::Exit");
		return hr;

	Error:

		hr = E_FAIL;
		NetOffice_ShimLoader_Analytics::WriteLog(L"ShimArguments::ReadShimSettings::FailExit");
		return hr;
	}

	HRESULT ShimArguments::ReadShimDefaults(MSXML::IXMLDOMNodePtr document)
	{
		NetOffice_ShimLoader_Analytics::WriteLog(L"ShimArguments::ReadShimDefaults::Enter");

		HRESULT hr = S_OK;
		MSXML::IXMLDOMNodePtr extensibilityDefaultNode = nullptr;
		MSXML::IXMLDOMNodePtr extensibilityFailNode = nullptr;

		extensibilityDefaultNode = document->selectSingleNode("/ShimLoader/Shim/Defaults/ExtensibilityDefaultResult");
		IfNullGo(extensibilityDefaultNode);
		extensibilityFailNode = document->selectSingleNode("/ShimLoader/Shim/Defaults/ExtensibilityFailResult");
		IfNullGo(extensibilityFailNode);

		EXTENSIBILITY_DEFAULT_RESULT = stol(extensibilityDefaultNode->text.copy());
		EXTENSIBILITY_FAIL_RESULT = stol(extensibilityFailNode->text.copy());

		NetOffice_ShimLoader_Analytics::WriteLog(L"ShimArguments::ReadShimDefaults::Exit");
		return hr;

	Error:

		hr = E_FAIL;
		NetOffice_ShimLoader_Analytics::WriteLog(L"ShimArguments::ReadShimDefaults::FailExit");
		return hr;
	}

	HRESULT ShimArguments::ReadManagedAddinAggregator(MSXML::IXMLDOMNodePtr document)
	{
		NetOffice_ShimLoader_Analytics::WriteLog(L"ShimArguments::ReadManagedAddinAggregator::Enter");

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

		IfFailGo(parser.ParseEx(folderPathNode->text, folderPathSubFolderNode->text, TargetManagedAggregator_Folder, MAX_PATH + 1, _documentPath));
		IfFailGo(parser.ParseEx(appDomainFolderPathNode->text, appDomainSubFolderNode->text, TargetManagedAggregator_AppDomain_BaseFolder, MAX_PATH + 1, _documentPath));

		lstrcpyn(TargetManagedAggregator_AssemblyName, assemblyNameNode->text, MAX_PATH + 1);
		lstrcpyn(TargetManagedAggregator_ClassName, classNameNode->text, MAX_PATH + 1);
		lstrcpyn(TargetManagedAggregator_AppDomain_FriendlyName, appDomainFriendlyNameNode->text, MAX_PATH + 1);
		lstrcpyn(Target_AssemblyName, targetAssemblyName->text, MAX_PATH + 1);
		lstrcpyn(Target_AssemblyFileName, targetAssemblyFileName->text, MAX_PATH + 1);
		lstrcpyn(Target_ConfigFileName, targetConfigFileName->text, MAX_PATH + 1);
		lstrcpyn(Target_ConnectClassName, targetClassName->text, MAX_PATH + 1);

		NetOffice_ShimLoader_Analytics::WriteLog(L"ShimArguments::ReadManagedAddinAggregator::Exit");
		return hr;

	Error:

		hr = E_FAIL;
		NetOffice_ShimLoader_Analytics::WriteLog(L"ShimArguments::ReadManagedAddinAggregator::FailExit");
		return hr;
	}

	HRESULT ShimArguments::ReadManagedUpdateAggregator(MSXML::IXMLDOMNodePtr document)
	{
		NetOffice_ShimLoader_Analytics::WriteLog(L"ShimArguments::ReadManagedUpdateAggregator::Enter");

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

		IfFailGo(parser.ParseEx(folderPathNode->text, folderPathSubFolderNode->text, UpdateManagedAggregator_Folder, MAX_PATH + 1, _documentPath));
		IfFailGo(parser.ParseEx(appDomainFolderPathNode->text, appDomainSubFolderNode->text, UpdateManagedAggregator_AppDomain_BaseFolder, MAX_PATH + 1, _documentPath));

		lstrcpyn(UpdateManagedAggregator_AssemblyName, assemblyNameNode->text, MAX_PATH + 1);
		lstrcpyn(UpdateManagedAggregator_ClassName, classNameNode->text, MAX_PATH + 1);
		lstrcpyn(UpdateManagedAggregator_AppDomain_FriendlyName, appDomainFriendlyNameNode->text, MAX_PATH + 1);
		lstrcpyn(Update_AssemblyName, targetAssemblyName->text, MAX_PATH + 1);
		lstrcpyn(Update_AssemblyFileName, targetAssemblyFileName->text, MAX_PATH + 1);
		lstrcpyn(Update_ConfigFileName, targetConfigFileName->text, MAX_PATH + 1);
		lstrcpyn(Update_ConnectClassName, targetClassName->text, MAX_PATH + 1);

		NetOffice_ShimLoader_Analytics::WriteLog(L"ShimArguments::ReadManagedUpdateAggregator::Exit");
		return hr;

	Error:

		hr = E_FAIL;
		NetOffice_ShimLoader_Analytics::WriteLog(L"ShimArguments::ReadManagedUpdateAggregator::FailExit");
		return hr;
	}

	HRESULT ShimArguments::ComInitialize()
	{
		HRESULT hr = S_OK;

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
