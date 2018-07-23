#pragma once
#include "stdafx.h"
#include "atlbase.h"
#include <fstream>
#include <iosfwd>
#include <iostream>
#include <ctime>
#include "Module.hpp"

extern WCHAR* LogFile_Path;
extern BOOL	  Internal_LogError_MessageBoxes_Enabled;

namespace NetOffice_ShimLoader_Analytics
{
	static const int _bufferSize = 1024;

	static HRESULT InitializeLog()
	{
		HRESULT hr = S_OK;

		WCHAR moduleFolderPath[MAX_PATH];
		WCHAR moduleFileName[MAX_PATH];
		WCHAR logFileName[MAX_PATH];

		hr = NetOffice_ShimLoader_Module::GetDllDirectory(moduleFolderPath, MAX_PATH);
		if (FAILED(hr))
			goto Error;
		hr = NetOffice_ShimLoader_Module::GetModuleFileName(moduleFileName, MAX_PATH);
		if (FAILED(hr))
			goto Error;
		PathStripPath(moduleFileName);

		StringCchCopy(logFileName, MAX_PATH, moduleFileName);
		StringCchCat(logFileName, MAX_PATH, L".log");

		PathAppend(LogFile_Path, moduleFolderPath);
		if (LogFile_Path && SUCCEEDED(PathAppend(LogFile_Path, logFileName)))
		{
			if (PathFileExists(LogFile_Path))
			{
				DeleteFile(LogFile_Path);
			}
		}

		return hr;

	Error:

		return hr;
	}

	static void _Write(LPCWSTR text)
	{
		std::wofstream myfile(LogFile_Path, std::ios::app);
		if (LogFile_Path && myfile.is_open())
		{
			myfile << text << std::endl;
			myfile.close();
		}
		else
		{
			#ifdef DEBUG
				MessageBox(GetDesktopWindow(), LogFile_Path, L"WriteLog::_Write::OpenFileError", 0);
			#endif
		}
	}

	static void WriteError(LPCWSTR text, ULONG value)
	{
		WCHAR buffer[_bufferSize];
		swprintf_s(buffer, _bufferSize, L"%s(%d)", text, value);
		_Write(text);

		#ifdef DEBUG

			if(Internal_LogError_MessageBoxes_Enabled)
				MessageBox(GetDesktopWindow(), buffer, L"WriteError", 0);

		#endif
	}

	static void WriteError(LPCWSTR text, HRESULT value)
	{
		WCHAR buffer[_bufferSize];
		swprintf_s(buffer, _bufferSize, L"%s(HR:%d)", text, value);
		_Write(text);

		#ifdef DEBUG

			if (Internal_LogError_MessageBoxes_Enabled)
				MessageBox(GetDesktopWindow(), buffer, L"WriteError", 0);

		#endif
	}

	static void WriteError(LPCWSTR text)
	{
		_Write(text);

		#ifdef DEBUG

				if (Internal_LogError_MessageBoxes_Enabled)
				MessageBox(GetDesktopWindow(), text, L"WriteError", 0);

		#endif
	}

	static void WriteError(LPCWSTR text, LPCWSTR text2)
	{
		WCHAR buffer[_bufferSize];
		swprintf_s(buffer, _bufferSize, L"%s(%s)", text, text2);
		_Write(text);

		#ifdef DEBUG

			if (Internal_LogError_MessageBoxes_Enabled)
				MessageBox(GetDesktopWindow(), buffer, L"WriteError", 0);

		#endif
	}

	static void WriteLog(LPCWSTR text, HRESULT value)
	{
		WCHAR buffer[_bufferSize];
		swprintf_s(buffer, _bufferSize, L"%s(HR:%d)", text, value);
		_Write(text);
	}

	static void WriteLog(LPCWSTR text)
	{
		_Write(text);
	}

	static void WriteLogTimeStamp(LPCWSTR text)
	{
		time_t t = time(NULL);
		struct tm buf;
		WCHAR timeBuffer[256];

		if (NULL == localtime_s(&buf, &t))
		{
			wcsftime(timeBuffer, 256, L"(%d-%m-%Y %I:%M:%S %p)", &buf);
			WCHAR buffer[_bufferSize];
			StringCchCopy(buffer, _bufferSize, text);
			StringCchCat(buffer, _bufferSize, timeBuffer);
			_Write(buffer);
		}
		else
		{
			_Write(text);
		}
	}

	static void WriteLog(LPCWSTR text, LPCWSTR text2)
	{
		WCHAR buffer[_bufferSize];
		swprintf_s(buffer, _bufferSize, L"%s(%s)", text, text2);
		_Write(text);
	}
}
