#pragma once
#include "stdafx.h"
#include "atlbase.h"
#include <fstream>
#include <iosfwd>
#include <iostream>
#include <ctime>
#include "Module.hpp"

extern WCHAR* LogFile_UnRegister_Path;
extern BOOL	  Internal_LogError_MessageBoxes_Enabled;

namespace NetOffice_ShimLoader_Analytics
{
	static HRESULT InitializeUnRegisterLog()
	{
		HRESULT hr = S_OK;

		WCHAR moduleFolderPath[MAX_PATH];
		WCHAR moduleFileName[MAX_PATH];
		WCHAR logUnregisterFileName[MAX_PATH];

		hr = NetOffice_ShimLoader_Module::GetDllDirectory(moduleFolderPath, MAX_PATH);
		if (FAILED(hr))
			goto Error;
		hr = NetOffice_ShimLoader_Module::GetModuleFileName(moduleFileName, MAX_PATH);
		if (FAILED(hr))
			goto Error;
		PathStripPath(moduleFileName);

		StringCchCopy(logUnregisterFileName, MAX_PATH, moduleFileName);
		StringCchCat(logUnregisterFileName, MAX_PATH, L".Unregister");
		StringCchCat(logUnregisterFileName, MAX_PATH, L".log");

		PathAppend(LogFile_UnRegister_Path, moduleFolderPath);
		if (LogFile_UnRegister_Path && SUCCEEDED(PathAppend(LogFile_UnRegister_Path, logUnregisterFileName)))
		{
			if (PathFileExists(LogFile_UnRegister_Path))
			{
				DeleteFile(LogFile_UnRegister_Path);
			}
		}

		return hr;

	Error:

		return hr;
	}

	static void _WriteUnRegister(LPCWSTR text)
	{
		std::wofstream myfile(LogFile_UnRegister_Path, std::ios::app);
		if (LogFile_UnRegister_Path && myfile.is_open())
		{
			myfile << text << std::endl;
			myfile.close();
		}
		else
		{
			#ifdef DEBUG
				MessageBox(GetDesktopWindow(), text, L"WriteLog::_WriteUnRegister::OpenFileError", 0);
			#endif
		}
	}

	static void WriteUnRegisterError(LPCWSTR text, HRESULT value)
	{
		WCHAR buffer[_bufferSize];
		swprintf_s(buffer, _bufferSize, L"%s(HR:%d)", text, value);
		_WriteUnRegister(text);

		#ifdef DEBUG

			if (Internal_LogError_MessageBoxes_Enabled)
				MessageBox(GetDesktopWindow(), buffer, L"WriteError", 0);

		#endif
	}

	static void WriteUnRegisterError(LPCWSTR text)
	{
		_WriteUnRegister(text);
	}

	static void WriteUnRegisterLog(LPCWSTR text)
	{
		_WriteUnRegister(text);
	}
}
