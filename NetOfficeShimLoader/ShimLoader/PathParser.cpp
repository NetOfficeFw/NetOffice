#include "stdafx.h"
#include "PathParser.h"
#include <strsafe.h>

namespace NetOffice_ShimLoader
{
	static HRESULT GetDllDirectory(TCHAR *szPath, DWORD nPathBufferSize);

	PathParser::PathParser()
	{
		_parseMap[L"LocalAppData"] = FOLDERID_LocalAppData;
		_parseMap[L"RoamingAppData"] = FOLDERID_RoamingAppData;
		_parseMap[L"CommonProgramData"] = FOLDERID_ProgramData;

		_parseLegacyMap[L"LocalAppData"] = CSIDL_LOCAL_APPDATA;
		_parseLegacyMap[L"RoamingAppData"] = CSIDL_APPDATA;
		_parseLegacyMap[L"CommonProgramData"] = CSIDL_COMMON_APPDATA;
	}

	PathParser::~PathParser()
	{
		_parseMap.clear();
		_parseLegacyMap.clear();
	}

	HRESULT PathParser::Parse(BSTR path, WCHAR* result, int maxLen, BSTR documentPath)
	{
		HRESULT hr = E_FAIL;

		if (OperatingSystemIsVistaOrAbove())
		{
			hr = ParseLegacyInternal(path, result, maxLen, documentPath);
		}
		else
		{
			hr = ParseInternal(path, result, maxLen, documentPath);
		}

		return hr;
	}

	HRESULT PathParser::ParseEx(BSTR path, BSTR subFolderPath, WCHAR* result, int maxLen, BSTR documentPath)
	{
		HRESULT hr = E_FAIL;

		if (OperatingSystemIsVistaOrAbove())
		{
			hr = ParseInternal(path, result, maxLen, documentPath);
			if (SUCCEEDED(hr) && 0 != wcscmp(subFolderPath, L""))
			{
				hr = PathAppend(result, subFolderPath);
				//StringCchCat(result, maxLen, subFolderPath);
			}
		}
		else
		{
			hr = ParseLegacyInternal(path, result, maxLen, documentPath);
			if (SUCCEEDED(hr) && 0 != wcscmp(subFolderPath, L""))
			{
				hr = PathAppend(result, subFolderPath);
				//StringCchCat(result, maxLen, subFolderPath);
			}
		}

		return hr;
	}

	HRESULT PathParser::ParseEx(BSTR path, BSTR subFolderPath, BSTR filePath, WCHAR* result, int maxLen)
	{
		return ParseEx(path, subFolderPath, filePath, result, maxLen, NULL);
	}

	HRESULT PathParser::ParseEx(BSTR path, BSTR subFolderPath, BSTR filePath, WCHAR* result, int maxLen, BSTR documentPath)
	{
		HRESULT hr = E_FAIL;

		if (OperatingSystemIsVistaOrAbove())
		{
			hr = ParseInternal(path, result, maxLen, documentPath);
			if (SUCCEEDED(hr) && 0 != wcscmp(subFolderPath, L""))
			{
				hr = PathAppend(result, subFolderPath);
				//StringCchCat(result, maxLen, subFolderPath);
			}
			if (SUCCEEDED(hr) && 0 != wcscmp(filePath, L""))
			{
				hr = PathAppend(result, filePath);
				//StringCchCat(result, maxLen, filePath);
			}
		}
		else
		{
			hr = ParseLegacyInternal(path, result, maxLen, documentPath);
			if (SUCCEEDED(hr) && 0 != wcscmp(subFolderPath, L""))
			{
				hr = PathAppend(result, subFolderPath);
				//StringCchCat(result, maxLen, subFolderPath);
			}
			if (SUCCEEDED(hr) && 0 != wcscmp(filePath, L""))
			{
				hr = PathAppend(result, filePath);
				//StringCchCat(result, maxLen, filePath);
			}
		}

		return hr;
	}

	GUID PathParser::FindGuid(BSTR path)
	{
		for (auto it = _parseMap.begin(); it != _parseMap.end(); it++)
		{
			auto first = it->first;
			auto folderId = it->second;
			if (0 == wcscmp(first, path))
				return folderId;
		}

		return GUID_NULL;
	}

	DWORD PathParser::FindDWord(BSTR path)
	{
		for (auto it = _parseLegacyMap.begin(); it != _parseLegacyMap.end(); it++)
		{
			auto first = it->first;
			auto folderId = it->second;
			if (0 == wcscmp(first, path))
				return folderId;
		}

		return 0;
	}

	HRESULT PathParser::ParseInternal(BSTR path, WCHAR* result, int maxLen, BSTR documentPath)
	{
		HRESULT hr = E_FAIL;
		PWSTR buffer = nullptr;
		auto folderId = FindGuid(path);

		if (NULL != documentPath && 0 == wcscmp(L"DocumentPath", path))
		{
			lstrcpyn(result, documentPath, maxLen);
		}
		else if (folderId != GUID_NULL)
		{
			hr = SHGetKnownFolderPath(folderId, KF_FLAG_DEFAULT, NULL, &buffer);
			if (SUCCEEDED(hr))
			{
				lstrcpyn(result, buffer, maxLen);
				StringCchCat(result, maxLen, L"\\");
			}
			if (path)
			{
				CoTaskMemFree(buffer);
				buffer = nullptr;
			}
		}
		else
		{
			hr = GetDllDirectory(result, maxLen);
		}

		return hr;
	}

	HRESULT PathParser::ParseLegacyInternal(BSTR path, WCHAR* result, int maxLen, BSTR documentPath)
	{
		HRESULT hr = E_FAIL;
		TCHAR* szPath = new TCHAR[MAX_PATH + 1];
		auto folderId = FindDWord(path);

		if (NULL != documentPath && 0 == wcscmp(L"DocumentPath", path))
		{
			lstrcpyn(result, documentPath, maxLen);
		}
		else if (folderId != 0)
		{
			hr = SHGetFolderPath(NULL, folderId, NULL, 0, szPath);
			if (SUCCEEDED(hr))
			{
				lstrcpyn(result, szPath, maxLen);
				StringCchCat(result, maxLen, L"\\");
			}
		}
		else
		{
			hr = GetDllDirectory(result, maxLen);
		}

		delete[] szPath;
		szPath = nullptr;

		return hr;
	}

	BOOL PathParser::OperatingSystemIsVistaOrAbove()
	{
		BOOL result;
		OSVERSIONINFOEX verex = { 0 };
		DWORD dwTypeMask = VER_MAJORVERSION | VER_MINORVERSION;
		DWORDLONG dwlConditionMask = 0;
		verex.dwOSVersionInfoSize = sizeof(OSVERSIONINFOEX);
		// https://docs.microsoft.com/de-de/windows/desktop/api/winnt/ns-winnt-_osversioninfoexa
		verex.dwMajorVersion = 6;
		verex.dwMinorVersion = 0;
		VER_SET_CONDITION(dwlConditionMask, VER_MAJORVERSION, VER_GREATER_EQUAL);
		VER_SET_CONDITION(dwlConditionMask, VER_MINORVERSION, VER_GREATER_EQUAL);
		result = VerifyVersionInfo(&verex, dwTypeMask, dwlConditionMask);
		return result;
	}

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
