#pragma once
#include "stdafx.h"
#include "atlbase.h"
#include <fstream>
#include <iosfwd>
#include <iostream>
#include <ctime>

namespace NetOffice_ShimLoader_Module
{
	static HRESULT GetModuleFileName(WCHAR* szPath, DWORD nPathBufferSize)
	{
		HMODULE hInstance = _AtlBaseModule.GetModuleInstance();
		if (0 == hInstance)
		{
			return E_FAIL;
		}

		DWORD dwFLen = ::GetModuleFileName(hInstance, szPath, nPathBufferSize);
		if (0 != dwFLen)
		{
			return S_OK;
		}
		else
		{
			return E_FAIL;
		}
	}

	static HRESULT GetDllDirectory(WCHAR* szPath, DWORD nPathBufferSize)
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