#include "StdAfx.h"
#include "ClrHost.h"
#include "Vars.h"

namespace NetOffice_ShimLoader
{
	using namespace mscorlib;
	using namespace NetOffice_Tools_Isolation;

	static HRESULT GetDllDirectory(TCHAR *szPath, DWORD nPathBufferSize);

	/***************************************************************************
	* Ctor, Dtor
	***************************************************************************/

	ClrHost::ClrHost(IShimHost* shimHost)
	{
		_runtimeHost = nullptr;
		_appDomain = nullptr;
		_outerAggregator = nullptr;
		_isLoaded = false;
		_lastLoadError = S_OK;
		_shimHost = shimHost;
		IncComponents(L"ClrHost");
	}

	ClrHost::~ClrHost()
	{
		Unload();
		DecComponents(L"ClrHost");
	}


	/***************************************************************************
	* ClrLoader Methods
	***************************************************************************/

	BOOL ClrHost::IsLoaded()
	{
		return _isLoaded;
	}

	OuterComAggregator* ClrHost::OuterAggregator()
	{
		return _outerAggregator;
	}

	HRESULT ClrHost::Load()
	{
		HRESULT hr = E_FAIL;

		CComVariant cvarManagedAggregator;

		_ObjectHandle* srpObjectHandle = nullptr;
		IOuterComAggregator* srpOuterComAggregator = nullptr;
		IManagedInnerComAggregator* srpManagedAggregator = nullptr;

		ICLRMetaHost* metaHost = nullptr;
		ICLRRuntimeInfo* runtimeInfo = nullptr;
		ICorRuntimeHost* runtimeHost = nullptr;

		auto directoryPath = TargetManagedAggregator_Folder;

		WCHAR fullInnerAddinFilePath[MAX_PATH + 1];
		IfFailGo(AppendPath(fullInnerAddinFilePath, directoryPath));
		IfFailGo(AppendPath(fullInnerAddinFilePath, Target_AssemblyFileName));

		WCHAR fullInnerAddinConfigFilePath[MAX_PATH + 1];
		IfFailGo(AppendPath(fullInnerAddinConfigFilePath, directoryPath));
		IfFailGo(AppendPath(fullInnerAddinConfigFilePath, Target_ConfigFileName));

		WCHAR runtimeVersion[30];
		DWORD cwchruntimeVersion = ARRAYSIZE(runtimeVersion);

		IUnknown* unkAppDomainSetup = nullptr;
		IAppDomainSetup* pDomainSetup = nullptr;
		IUnknown* unkAppDomain = nullptr;

		IfFailGo(CLRCreateInstance(CLSID_CLRMetaHost, IID_ICLRMetaHost, (void**)&metaHost));
		IfFailGo(metaHost->GetVersionFromFile(fullInnerAddinFilePath, runtimeVersion, &cwchruntimeVersion));
		IfFailGo(metaHost->GetRuntime(runtimeVersion, IID_ICLRRuntimeInfo, (void**)&runtimeInfo));
		IfFailGo(runtimeInfo->SetDefaultStartupFlags(STARTUP_LOADER_OPTIMIZATION_MULTI_DOMAIN_HOST, NULL));
		IfFailGo(runtimeInfo->GetInterface(CLSID_CorRuntimeHost, IID_ICorRuntimeHost, (void**)&runtimeHost));

		IfFailGo(runtimeHost->CreateDomainSetup(&unkAppDomainSetup));
		IfFailGo(unkAppDomainSetup->QueryInterface(__uuidof(pDomainSetup), (LPVOID*)&pDomainSetup));
		pDomainSetup->put_ApplicationBase(CComBSTR(TargetManagedAggregator_AppDomain_BaseFolder));

		if (PathFileExists(fullInnerAddinConfigFilePath))
		{
			IfFailGo(pDomainSetup->put_ConfigurationFile(fullInnerAddinConfigFilePath));
		}

		auto appDomainFriendlyName = wcslen(TargetManagedAggregator_AppDomain_FriendlyName) > 0 ? T2W(TargetManagedAggregator_AppDomain_FriendlyName) : T2W(directoryPath);
		IfFailGo(runtimeHost->CreateDomainEx(appDomainFriendlyName, pDomainSetup, 0, &unkAppDomain));
		IfFailGo(unkAppDomain->QueryInterface(__uuidof(_appDomain), (LPVOID*)&_appDomain));
		IfFailGo(_appDomain->CreateInstance(
			CComBSTR(TargetManagedAggregator_AssemblyName),
			CComBSTR(TargetManagedAggregator_ClassName),
			&srpObjectHandle));

		_outerAggregator = new OuterComAggregator();

		IfFailGo(srpObjectHandle->Unwrap(&cvarManagedAggregator));
		IfFailGo(cvarManagedAggregator.pdispVal->QueryInterface(&srpManagedAggregator));
		IfFailGo(_outerAggregator->QueryInterface(IID_IOuterComAggregator, (LPVOID*)&srpOuterComAggregator));

		IfFailGo(srpManagedAggregator->CreateAggregatedInstance(
			CComBSTR(Target_AssemblyName),
			CComBSTR(Target_ConnectClassName),
			srpOuterComAggregator, _shimHost));

		_runtimeHost = runtimeHost;

		_isLoaded = true;
		return S_OK;

	Error:

		_lastLoadError = hr;
		_runtimeHost = runtimeHost;
		Unload();
		return hr;
	}

	HRESULT ClrHost::Unload()
	{
		HRESULT hr = S_OK;
		IUnknown* pUnkDomain = NULL;

		if (_outerAggregator)
		{
			delete _outerAggregator;
			_outerAggregator = nullptr;
		}

		if (_appDomain)
		{
			if (_runtimeHost)
			{
				hr = _appDomain->QueryInterface(__uuidof(IUnknown), (LPVOID*)&pUnkDomain);
				if (SUCCEEDED(hr))
					hr = _runtimeHost->UnloadDomain(pUnkDomain);
				_runtimeHost = nullptr;
			}

			_appDomain->Release();
			_appDomain = NULL;
		}

		if (_runtimeHost)
		{
			_runtimeHost->Release();
			_runtimeHost = nullptr;
		}

		if (pUnkDomain)
		{
			pUnkDomain->Release();
			pUnkDomain = nullptr;
		}

		_isLoaded = false;

		return hr;
	}

	HRESULT ClrHost::LastLoadError()
	{
		return _lastLoadError;
	}

	HRESULT ClrHost::AppendPath(LPWSTR pszPath, LPCWSTR pszMore)
	{
		HRESULT hr = S_OK;
		if (!PathAppend(pszPath, pszMore))
		{
			hr = E_UNEXPECTED;
		}
		return hr;
	}


	/***************************************************************************
	* Static Utils
	***************************************************************************/

	static HRESULT GetDllDirectory(TCHAR* szPath, DWORD nPathBufferSize)
	{
		HMODULE hInstance = _AtlBaseModule.GetModuleInstance();
		if (0 == hInstance)
		{
			return E_FAIL;
		}

		TCHAR szModule[MAX_PATH + 1];
		DWORD dwFLen = ::GetModuleFileName(hInstance, szModule, MAX_PATH);
		if (0 == dwFLen)
		{
			return E_FAIL;
		}

		TCHAR* pszFileName;
		dwFLen = ::GetFullPathName(szModule, nPathBufferSize, szPath, &pszFileName);
		if (0 == dwFLen || dwFLen >= nPathBufferSize)
		{
			return E_FAIL;
		}

		*pszFileName = 0;
		return S_OK;
	}
}
