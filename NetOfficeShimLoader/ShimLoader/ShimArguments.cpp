#include "stdafx.h"
#include "ShimArguments.h"
#include "string.h"
#include "wchar.h"
#include <fstream>
#include <iostream>
#include <string>

static HRESULT GetDllDirectory(TCHAR *szPath, DWORD nPathBufferSize);

using namespace std;

ShimArguments::ShimArguments()
{
	_components++;
}

ShimArguments::~ShimArguments()
{
	_components--;
}

HRESULT ShimArguments::Load()
{
	HRESULT hr = E_FAIL;
	bool b = FALSE;
	std::ifstream infile;
	std::string line;

	WCHAR directoryPath[MAX_PATH + 1];
	IfFailGo(GetDllDirectory(directoryPath, ARRAYSIZE(directoryPath)));

	WCHAR moduleFileName[MAX_PATH + 1];
	IfFailGo(GetModuleFileName(_module, moduleFileName, ARRAYSIZE(moduleFileName)));

	WCHAR fullSettingsFilePath[MAX_PATH + 1];
	IfFailGo(AppendPath(fullSettingsFilePath, directoryPath));
	IfFailGo(AppendPath(fullSettingsFilePath, moduleFileName));
	PWSTR target = StrCatBuff(fullSettingsFilePath, L".ShimSettings", ARRAYSIZE(fullSettingsFilePath));

	IfFalseGo(PathFileExists(target));

	infile = std::ifstream(target);

	while (std::getline(infile, line))
	{
		// reading lines here and parsing values
		// to replace Vars.hpp
	}

	hr = S_OK;

Error:

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