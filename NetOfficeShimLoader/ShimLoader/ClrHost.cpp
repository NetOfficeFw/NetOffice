#include "StdAfx.h"
#include "ClrHost.h"

using namespace mscorlib;
using namespace NetOffice_Tools_Isolation;

static HRESULT GetDllDirectory(TCHAR *szPath, DWORD nPathBufferSize);

/***************************************************************************
* Ctor, Dtor
***************************************************************************/

ClrHost::ClrHost() : _runtimeHost(NULL), _appDomain(NULL), _aggregator(NULL)
{
	_refCounter = 0;
	_isLoaded = false;
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

bool ClrHost::IsLoaded()
{
	return _isLoaded;
}

OuterComAggregator* ClrHost::Aggregator()
{
	return _aggregator;
}

HRESULT ClrHost::Load()
{
	HRESULT hr = E_FAIL;

	_ObjectHandle* srpObjectHandle;
	CComVariant cvarManagedAggregator;
	IOuterComAggregator* srpComAggregator;
	IManagedInnerAggregator* srpManagedAggregator;

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

	_aggregator = new OuterComAggregator();

	IfFailGo(srpObjectHandle->Unwrap(&cvarManagedAggregator));
	IfFailGo(cvarManagedAggregator.pdispVal->QueryInterface(&srpManagedAggregator));
	IfFailGo(_aggregator->QueryInterface(__uuidof(IOuterComAggregator), (LPVOID*)&srpComAggregator));
	IfFailGo(srpManagedAggregator->CreateAggregatedInstance(CComBSTR(Target_AssemblyName), CComBSTR(Target_ConnectClassName), srpComAggregator));

	_runtimeHost = runtimeHost;

	_isLoaded = true;
	return S_OK;

Error:

	_runtimeHost = runtimeHost;
	Unload();
	return hr;
}

HRESULT ClrHost::Unload()
{
	HRESULT hr = S_OK;
	IUnknown* pUnkDomain = NULL;

	if (_aggregator)
	{
		delete _aggregator;
		_aggregator = nullptr;
	}

	if (_appDomain)
	{
		if (_runtimeHost)
		{
			hr = _appDomain->QueryInterface(__uuidof(IUnknown), (LPVOID*)&pUnkDomain);
			if(S_OK == hr)
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

	_isLoaded = true;

	return hr;
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
