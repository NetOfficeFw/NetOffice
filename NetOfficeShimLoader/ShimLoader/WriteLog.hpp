#pragma once
#include "stdafx.h"
#include "atlbase.h"
#include <fstream>
#include <iosfwd>
#include <iostream>

namespace NetOffice_ShimLoader_Analytics
{
	static WCHAR _logFilePath[MAX_PATH + 1];

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

	static void InitializeLog()
	{
		if (SUCCEEDED(GetDllDirectory(_logFilePath, MAX_PATH))
			&& SUCCEEDED(PathAppend(_logFilePath, L"ShimLoader.txt")))
		{
			if (PathFileExists(_logFilePath))
			{
				DeleteFile(_logFilePath);
			}
		}
	}

	static void dump(LPCWSTR text)
	{
		std::wofstream myfile("C:\\myfile.txt", std::ios::app);
		if (myfile.is_open())
		{
			myfile << text << std::endl;
			myfile.close();
		}
	}

	static void WriteError(LPCWSTR text, ULONG value)
	{
		MessageBox(GetDesktopWindow(), text, L"WriteError", 0);
	}

	static void WriteError(LPCWSTR text, HRESULT hr)
	{
		MessageBox(GetDesktopWindow(), text, L"WriteError", 0);
	}

	static void WriteError(LPCWSTR text)
	{
		MessageBox(GetDesktopWindow(), text, L"WriteError", 0);
	}

	static void WriteError(LPCWSTR text, LPCWSTR text2)
	{
		//MessageBox(GetDesktopWindow(), text, L"WriteLog", 0);
	}
	static void WriteLog(LPCWSTR text, HRESULT hr)
	{
		//MessageBox(GetDesktopWindow(), text, L"WriteLog", 0);
	}
	static void WriteLog(LPCWSTR text)
	{
		//MessageBox(GetDesktopWindow(), text, L"WriteLog", 0);
	}
	static void WriteLog(LPCWSTR text, LPCWSTR text2)
	{
		//MessageBox(GetDesktopWindow(), text, L"WriteLog", 0);
	}

	//static void WriteLog(const char* szString)
	//{
	//	//MessageBox(GetDesktopWindow(), szString, L"WriteLog", 0);

	//	//ShimDebugMessageBox(L"", L"");

	//	/*FILE* file = nullptr;
	//	auto result = _wfopen_s(&file, _logFilePath, L"a");
	//	if (NULL != result)
	//	{
	//		fprintf(file, "%s\n", szString);
	//		fclose(file);
	//	}*/
	//}
}
