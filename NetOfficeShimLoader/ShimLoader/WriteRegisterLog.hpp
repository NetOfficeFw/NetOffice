#pragma once
#include "stdafx.h"
#include "atlbase.h"
#include <fstream>
#include <iosfwd>
#include <iostream>
#include <ctime>
#include "Module.hpp"

extern WCHAR* LogFile_Register_Path;
extern BOOL	  Internal_LogError_MessageBoxes_Enabled;

namespace NetOffice_ShimLoader_Analytics
{
	static HRESULT InitializeWriteRegisterLog()
	{
		HRESULT hr = S_OK;

		WCHAR moduleFolderPath[MAX_PATH];
		WCHAR moduleFileName[MAX_PATH];
		WCHAR logRegisterFileName[MAX_PATH];

		hr = NetOffice_ShimLoader_Module::GetDllDirectory(moduleFolderPath, MAX_PATH);
		if (FAILED(hr))
			goto Error;
		hr = NetOffice_ShimLoader_Module::GetModuleFileName(moduleFileName, MAX_PATH);
		if (FAILED(hr))
			goto Error;
		PathStripPath(moduleFileName);

		StringCchCopy(logRegisterFileName, MAX_PATH, moduleFileName);
		StringCchCat(logRegisterFileName, MAX_PATH, L".Register");
		StringCchCat(logRegisterFileName, MAX_PATH, L".log");

		PathAppend(LogFile_Register_Path, moduleFolderPath);
		if (LogFile_Register_Path && SUCCEEDED(PathAppend(LogFile_Register_Path, logRegisterFileName)))
		{
			if (PathFileExists(LogFile_Register_Path))
			{
				DeleteFile(LogFile_Register_Path);
			}
		}

		return hr;

	Error:

		return hr;
	}

	static void _WriteRegister(LPCWSTR text)
	{
		std::wofstream myfile(LogFile_Register_Path, std::ios::app);
		if (LogFile_Register_Path && myfile.is_open())
		{
			myfile << text << std::endl;
			myfile.close();
		}
		else
		{
			#ifdef DEBUG
				MessageBox(GetDesktopWindow(), text, L"WriteLog::_WriteRegister::OpenFileError", 0);
			#endif
		}
	}

	static void WriteRegisterError(LPCWSTR text, HRESULT value)
	{
		WCHAR buffer[_bufferSize];
		swprintf_s(buffer, _bufferSize, L"%s(HR:%d)", text, value);
		_WriteRegister(text);

		#ifdef DEBUG

			if (Internal_LogError_MessageBoxes_Enabled)
				MessageBox(GetDesktopWindow(), buffer, L"WriteError", 0);

		#endif
	}

	static void WriteRegisterError(LPCWSTR text)
	{
		_WriteRegister(text);
	}

	static void WriteRegisterLog(LPCWSTR text)
	{
		_WriteRegister(text);
	}

}
