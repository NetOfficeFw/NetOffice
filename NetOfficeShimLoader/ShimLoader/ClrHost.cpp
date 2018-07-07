#include "StdAfx.h"
#include "ClrHost.h"

namespace NetOffice_ShimLoader
{
	using namespace mscorlib;
	using namespace NetOffice_Tools_Isolation;

	static HRESULT GetDllDirectory(TCHAR *szPath, DWORD nPathBufferSize);

	/***************************************************************************
	* Ctor, Dtor
	***************************************************************************/

	ClrHost::ClrHost(IOuterUpdateAggregator* updateAggregator) : _runtimeHost(NULL), _appDomain(NULL), _outerAggregator(NULL)
	{
		_isLoaded = false;
		_lastLoadError = S_OK;
		_updateAggregator = updateAggregator;
		_components++;
	}

	ClrHost::~ClrHost()
	{
		Unload();
		_components--;
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
		IManagedInnerAggregator* srpManagedAggregator = nullptr;

		ICLRMetaHost* metaHost = nullptr;
		ICLRRuntimeInfo* runtimeInfo = nullptr;
		ICorRuntimeHost* runtimeHost = nullptr;

		WCHAR directoryPath[MAX_PATH + 1];
		IfFailGo(GetDllDirectory(directoryPath, ARRAYSIZE(directoryPath)));

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
		pDomainSetup->put_ApplicationBase(CComBSTR(directoryPath));

		if (PathFileExists(fullInnerAddinConfigFilePath))
		{
			IfFailGo(pDomainSetup->put_ConfigurationFile(fullInnerAddinConfigFilePath));
		}

		IfFailGo(runtimeHost->CreateDomainEx(T2W(directoryPath), pDomainSetup, 0, &unkAppDomain));
		IfFailGo(unkAppDomain->QueryInterface(__uuidof(_appDomain), (LPVOID*)&_appDomain));
		IfFailGo(_appDomain->CreateInstance(
			CComBSTR(ManagedAggregator_AssemblyName),
			CComBSTR(ManagedAggregator_ClassName),
			&srpObjectHandle));

		_outerAggregator = new OuterComAggregator();

		IfFailGo(srpObjectHandle->Unwrap(&cvarManagedAggregator));
		IfFailGo(cvarManagedAggregator.pdispVal->QueryInterface(&srpManagedAggregator));
		IfFailGo(_outerAggregator->QueryInterface(__uuidof(IOuterComAggregator), (LPVOID*)&srpOuterComAggregator));

		IfFailGo(srpManagedAggregator->CreateAggregatedInstance(
			CComBSTR(Target_AssemblyName),
			CComBSTR(Target_ConnectClassName),
			srpOuterComAggregator, _updateAggregator));

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

	static HRESULT GetDllDirectory(TCHAR *szPath, DWORD nPathBufferSize)
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
