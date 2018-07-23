#include "stdafx.h"
#include "DllRegisterMode.hpp"
#include "Vars.h"

using namespace NetOffice_ShimLoader;

namespace NetOffice_ShimLoader_Register32On64
{
	DWORD _regKeyOptions = KEY_ALL_ACCESS;

	HRESULT RegisterCOMComponent(HINSTANCE module, LPCWSTR progId, LPCWSTR classId, LPCWSTR version, LPCWSTR description, RegisterMode mode);
	HRESULT UnregisterCOMComponent(LPCWSTR progId, LPCWSTR classId, LPCWSTR version, RegisterMode mode);
	HRESULT RegisterCOMAddin(LPCWSTR pszOfficeApp, LPCWSTR pszProgID, LPCWSTR pszFriendlyName, LPCWSTR pszDescription, DWORD dwStartupContext, DWORD dwCommandLineSafe, bool registerPerMachine);
	HRESULT UnRegisterCOMAddin(LPCWSTR pszOfficeApp, LPCWSTR pszProgID, bool registerPerMachine);

	HKEY TargetRootKey(RegisterMode mode);
	void ProgIdSubKey(LPCWSTR progId, RegisterMode mode, WCHAR* progIdKey, int maxLen);
	void ClassIdSubKey(LPCWSTR classId, RegisterMode mode, WCHAR* classIdKey, int maxLen);
	HRESULT SetCustomValue(HKEY hKey, PCustomRegisterValue value);
	BOOL SetKeyAndValue(HKEY hKeyRoot, LPCWSTR pszPath, LPCWSTR pszSubkey1, LPCWSTR pszSubkey2, LPCWSTR pszSubkey3, LPCWSTR pszvalueName, LPCWSTR pszValue);
	LONG RecursiveDeleteKey(HKEY hKeyParent, LPCWSTR pszKeyChild);
	LONG RecursiveDeleteKey(HKEY hKeyParent, LPCWSTR pszKeyChild, LPCWSTR pszKeyChild2);
	LONG RecursiveDeleteKey(HKEY hKeyParent, LPCWSTR pszKeyChild, LPCWSTR pszKeyChild2, LPCWSTR pszKeyChild3);

	HRESULT DllRegister(HINSTANCE module, LPCWSTR officeApplications[], DWORD addinLoadBehavior, DWORD addinCommandLineSafe, WCHAR* progId, WCHAR* classId, WCHAR* friendlyName, WCHAR* description, WCHAR* version, RegisterMode mode, BOOL addinRegistration)
	{
		NetOffice_ShimLoader_Analytics::WriteRegisterLog(L"Register32On64::DllRegister::Enter");

		HRESULT hr = S_OK;

		if (NULL == module || NULL == officeApplications ||
			!progId || !progId[0] || !classId || !classId[0] ||
			!friendlyName || !friendlyName[0] || !description || !description[0] ||
			!version || !version[0])
		{
			hr = E_INVALIDARG;
			NetOffice_ShimLoader_Analytics::WriteRegisterError(L"Register32On64::DllRegister::FailExit", hr);
			return hr;
		}

		hr = RegisterCOMComponent(module, progId, classId, version, description, mode);
		if (SUCCEEDED(hr) && addinRegistration)
		{
			size_t arraySize = ShimProxy_Host_Application_Length;
			for (size_t i = 0; i < arraySize; i++)
			{
				hr = RegisterCOMAddin(officeApplications[i], progId, friendlyName, description, addinLoadBehavior, addinCommandLineSafe, System == mode);
				if (!SUCCEEDED(hr))
				{
					NetOffice_ShimLoader_Analytics::WriteRegisterError(L"Register32On64::DllRegister::Error#RegisterCOMAddin", hr);
					break;
				}
			}
		}
		else
		{
			NetOffice_ShimLoader_Analytics::WriteRegisterError(L"Register32On64::DllRegister::Error#RegisterCOMComponent", hr);
		}

		if (SUCCEEDED(hr))
			NetOffice_ShimLoader_Analytics::WriteRegisterLog(L"Register32On64::DllRegister::Exit");
		else
			NetOffice_ShimLoader_Analytics::WriteRegisterError(L"Register32On64::DllRegister::FailExit", hr);

		return hr;
	}

	HRESULT DllUnregister(LPCWSTR officeApplications[], WCHAR* progId, WCHAR* classId, WCHAR* version, RegisterMode mode, BOOL addinRegistration)
	{
		NetOffice_ShimLoader_Analytics::WriteUnRegisterLog(L"Register32On64::DllUnregister::Enter");

		HRESULT hr = S_OK;
		HRESULT addin = S_OK;

		if (NULL == officeApplications || !progId || !progId[0] || !classId || !classId[0] || !version || !version[0])
		{
			hr = E_INVALIDARG;
			NetOffice_ShimLoader_Analytics::WriteUnRegisterError(L"Register32On64::DllUnregister::FailExit", hr);
			return hr;
		}

		if (addinRegistration)
		{
			size_t arraySize = ShimProxy_Host_Application_Length;
			for (size_t i = 0; i < arraySize; i++)
			{
				if (!SUCCEEDED(UnRegisterCOMAddin(officeApplications[i], progId, 0 == mode)))
				{
					NetOffice_ShimLoader_Analytics::WriteUnRegisterError(L"Register32On64::DllUnregister::Error#UnRegisterCOMAddin", hr);
					addin = E_FAIL;
				}
			}
		}

		hr = UnregisterCOMComponent(progId, classId, version, mode);
		if(FAILED(hr))
			NetOffice_ShimLoader_Analytics::WriteUnRegisterError(L"Register32On64::DllUnregister::Error#UnregisterCOMComponent", hr);

		if(SUCCEEDED(addin != S_OK ? addin : hr))
			NetOffice_ShimLoader_Analytics::WriteUnRegisterLog(L"Register32On64::DllUnregister::Exit");
		else
			NetOffice_ShimLoader_Analytics::WriteUnRegisterError(L"Register32On64::DllUnregister::FailExit", (addin != S_OK ? addin : hr));

		return addin != S_OK ? addin : hr;
	}

	HRESULT RegisterCOMComponent(HINSTANCE module, LPCWSTR progId, LPCWSTR classId, LPCWSTR version, LPCWSTR description, RegisterMode mode)
	{
		NetOffice_ShimLoader_Analytics::WriteRegisterLog(L"Register32On64::RegisterCOMComponent::Enter");

		HRESULT hr = S_OK;
		BOOL setKeyResult = FALSE;

		WCHAR moduleFullFileName[512];
		DWORD dwResult = ::GetModuleFileName(module, moduleFullFileName, 512);
		if (0 != dwResult)
		{
			HKEY targetRootKey = TargetRootKey(mode);

			WCHAR classIdKey[512];
			ClassIdSubKey(classId, mode, classIdKey, 512);

			WCHAR progIdKey[512];
			ProgIdSubKey(progId, mode, progIdKey, 512);

			// Target Key ProgId
			setKeyResult = SetKeyAndValue(targetRootKey, progIdKey, NULL, NULL, NULL, NULL, progId);
			IfFalseGo(setKeyResult);
			setKeyResult = SetKeyAndValue(targetRootKey, progIdKey, L"CLSID", NULL, NULL, NULL, classId);
			IfFalseGo(setKeyResult);

			// Target Key IID
			setKeyResult = SetKeyAndValue(targetRootKey, classIdKey, NULL, NULL, NULL, NULL, progId);
			IfFalseGo(setKeyResult);
			setKeyResult = SetKeyAndValue(targetRootKey, classIdKey, L"InprocServer32", NULL, NULL, L"ThreadingModel", L"Apartment");
			IfFalseGo(setKeyResult);
			setKeyResult = SetKeyAndValue(targetRootKey, classIdKey, L"InprocServer32", NULL, NULL, NULL, moduleFullFileName);
			IfFalseGo(setKeyResult);
			setKeyResult = SetKeyAndValue(targetRootKey, classIdKey, L"InprocServer32", version, NULL, L"ThreadingModel", L"Apartment");
			IfFalseGo(setKeyResult);
			setKeyResult = SetKeyAndValue(targetRootKey, classIdKey, L"InprocServer32", version, NULL, NULL, moduleFullFileName);
			IfFalseGo(setKeyResult);
			setKeyResult = SetKeyAndValue(targetRootKey, classIdKey, L"ProgId", NULL, NULL, NULL, progId);
			IfFalseGo(setKeyResult);

			//if (mode != User)
			//{
			//	// HKEY_CLASSES_ROOT ProgId
			//	setKeyResult = SetKeyAndValue(HKEY_CLASSES_ROOT, progId, NULL, NULL, NULL, NULL, progId);
			//	IfFalseGo(setKeyResult);
			//	setKeyResult = SetKeyAndValue(HKEY_CLASSES_ROOT, progId, L"CLSID", NULL, NULL, NULL, classId);
			//	IfFalseGo(setKeyResult);

			//	// HKEY_CLASSES_ROOT IID
			//	setKeyResult = SetKeyAndValue(HKEY_CLASSES_ROOT, L"CLSID", classId, NULL, NULL, NULL, progId);
			//	IfFalseGo(setKeyResult);
			//	setKeyResult = SetKeyAndValue(HKEY_CLASSES_ROOT, L"CLSID", classId, L"InprocServer32", NULL, L"ThreadingModel", L"Apartment");
			//	IfFalseGo(setKeyResult);
			//	setKeyResult = SetKeyAndValue(HKEY_CLASSES_ROOT, L"CLSID", classId, L"InprocServer32", NULL, NULL, moduleFullFileName);
			//	IfFalseGo(setKeyResult);
			//	setKeyResult = SetKeyAndValue(HKEY_CLASSES_ROOT, L"CLSID", classId, L"InprocServer32", version, L"ThreadingModel", L"Apartment");
			//	IfFalseGo(setKeyResult);
			//	setKeyResult = SetKeyAndValue(HKEY_CLASSES_ROOT, L"CLSID", classId, L"InprocServer32", version, NULL, moduleFullFileName);
			//	IfFalseGo(setKeyResult);
			//	setKeyResult = SetKeyAndValue(HKEY_CLASSES_ROOT, L"CLSID", classId, L"ProgId", NULL, NULL, progId);
			//	IfFalseGo(setKeyResult);
			//}
		}
		else
		{
			hr = E_FAIL;
			goto Error;
		}

		NetOffice_ShimLoader_Analytics::WriteRegisterLog(L"Register32On64::RegisterCOMComponent::Exit");
		return hr;

	Error:

		NetOffice_ShimLoader_Analytics::WriteRegisterError(L"Register32On64::RegisterCOMComponent::FailExit", hr);
		return hr;
	}

	HRESULT UnregisterCOMComponent(LPCWSTR progId, LPCWSTR classId, LPCWSTR version, RegisterMode mode)
	{
		NetOffice_ShimLoader_Analytics::WriteUnRegisterLog(L"Register32On64::UnregisterCOMComponent::Enter");

		HRESULT hr = S_OK;
		LONG deleteKeyResult = 0;

		HKEY hKeyRoot = TargetRootKey(mode);
		WCHAR classIdKey[512];
		ClassIdSubKey(classId, mode, classIdKey, 512);
		WCHAR progIdKey[512];
		ProgIdSubKey(progId, mode, progIdKey, 512);

		//if (mode != User)
		//{
		//	deleteKeyResult = RecursiveDeleteKey(HKEY_CLASSES_ROOT, progId);
		//	IfZeroGo(deleteKeyResult);

		//	deleteKeyResult = RecursiveDeleteKey(HKEY_CLASSES_ROOT, L"CLSID", classId);
		//	IfZeroGo(deleteKeyResult);
		//}

		deleteKeyResult = RecursiveDeleteKey(hKeyRoot, progIdKey);
		IfNotZeroGo(deleteKeyResult);

		deleteKeyResult = RecursiveDeleteKey(hKeyRoot, classIdKey);
		IfNotZeroGo(deleteKeyResult);

		NetOffice_ShimLoader_Analytics::WriteUnRegisterLog(L"Register32On64::UnregisterCOMComponent::Exit");
		return hr;

	Error:

		NetOffice_ShimLoader_Analytics::WriteUnRegisterError(L"Register32On64::UnregisterCOMComponent::FailExit", hr);
		return hr;
	}

	HRESULT RegisterCOMAddin(LPCWSTR pszOfficeApp, LPCWSTR pszProgID, LPCWSTR pszFriendlyName, LPCWSTR pszDescription, DWORD dwStartupContext, DWORD dwCommandLineSafe, bool registerPerMachine)
	{
		NetOffice_ShimLoader_Analytics::WriteRegisterLog(L"Register32On64::RegisterCOMAddin::Enter");

		HRESULT hr = S_OK;
		WCHAR szKeyBuf[1024];
		DWORD dwTemp = 0;
		bool keyCreated = false;
		HKEY hKey;

		StringCchCopy(szKeyBuf, 1024, L"Software\\Microsoft\\Office\\");
		StringCchCat(szKeyBuf, 1024, pszOfficeApp);
		StringCchCat(szKeyBuf, 1024, L"\\Addins\\");
		StringCchCat(szKeyBuf, 1024, pszProgID);

		HKEY root = registerPerMachine ? HKEY_LOCAL_MACHINE : HKEY_CURRENT_USER;
		IfFailGo(RegCreateKeyEx(root, szKeyBuf, 0, NULL, REG_OPTION_NON_VOLATILE, _regKeyOptions, NULL, &hKey, NULL));

		IfFailGo(RegSetValueEx(hKey, L"LoadBehavior", 0, REG_DWORD, (BYTE*)&dwStartupContext, sizeof(DWORD)));
		IfFailGo(RegSetValueEx(hKey, L"CommandLineSafe", 0, REG_DWORD, (BYTE*)&dwCommandLineSafe, sizeof(DWORD)));

		if (NULL != pszFriendlyName)
		{
#if UNICODE
			dwTemp = lstrlen(pszFriendlyName) * 2 + 2;
#else
			dwTemp = lstrlen(pszFriendlyName) + 1;
#endif
			IfFailGo(RegSetValueEx(hKey, L"FriendlyName", 0, REG_SZ, (BYTE*)pszFriendlyName, dwTemp));

#if UNICODE
			dwTemp = lstrlen(pszDescription) * 2 + 2;
#else
			dwTemp = lstrlen(pszDescription) + 1;
#endif
			IfFailGo(RegSetValueEx(hKey, L"Description", 0, REG_SZ, (BYTE*)pszDescription, dwTemp));
		}

		if (NULL != Custom_Register_Values)
		{
			size_t arraySize = Custom_Register_Values_Length;
			for (size_t i = 0; i < arraySize; i++)
			{
				auto value = Custom_Register_Values[i];
				if (value->SeemsToBeValid())
					SetCustomValue(hKey, value);
			}
		}

		RegCloseKey(hKey);
		NetOffice_ShimLoader_Analytics::WriteRegisterLog(L"Register32On64::RegisterCOMAddin::Exit");
		return hr;

	Error:

		if (keyCreated)
		{
			RegCloseKey(hKey);
			RegDeleteKey(hKey, szKeyBuf);
		}
		NetOffice_ShimLoader_Analytics::WriteRegisterError(L"Register32On64::RegisterCOMAddin::FailExit", hr);
		return hr;
	}

	HRESULT UnRegisterCOMAddin(LPCWSTR pszOfficeApp, LPCWSTR pszProgID, bool registerPerMachine)
	{
		NetOffice_ShimLoader_Analytics::WriteUnRegisterLog(L"Register32On64::UnRegisterCOMAddin::Enter");

		HRESULT hr = S_OK;
		HRESULT result = S_OK;

		HKEY root = registerPerMachine ? HKEY_LOCAL_MACHINE : HKEY_CURRENT_USER;
		WCHAR szKeyBuf[1024];
		StringCchCopy(szKeyBuf, 1024, L"Software\\Microsoft\\Office\\");
		StringCchCat(szKeyBuf, 1024, pszOfficeApp);
		StringCchCat(szKeyBuf, 1024, L"\\Addins\\");
		StringCchCat(szKeyBuf, 1024, pszProgID);

		hr = RecursiveDeleteKey(root, szKeyBuf);
		if(SUCCEEDED(hr))
			NetOffice_ShimLoader_Analytics::WriteUnRegisterLog(L"Register32On64::UnRegisterCOMAddin::Exit");
		else
			NetOffice_ShimLoader_Analytics::WriteUnRegisterError(L"Register32On64::UnRegisterCOMAddin::FailExit", hr);
		return result;
	}

	HRESULT SetCustomValue(HKEY hKey, PCustomRegisterValue value)
	{
		WCHAR valueBuffer[256];
		value->ProcessedValue(valueBuffer, 256);

		DWORD dwTemp = 0;
#if UNICODE
		dwTemp = lstrlen(valueBuffer) * 2 + 2;
#else
		dwTemp = lstrlen(valueBuffer) + 1;
#endif

		return RegSetValueEx(hKey, value->Name(), 0, value->RegKind(), (BYTE*)valueBuffer, dwTemp);
	}

	HKEY TargetRootKey(RegisterMode mode)
	{
		HKEY hKeyRoot = HKEY_CURRENT_USER;
		switch (mode)
		{
		case System:
		case SystemComponentAndUserAddin:
			hKeyRoot = HKEY_LOCAL_MACHINE;
			break;
		}
		return hKeyRoot;
	}

	void ClassIdSubKey(LPCWSTR classId, RegisterMode mode, WCHAR* classIdKey, int maxLen)
	{
		StringCchCopy(classIdKey, maxLen, L"Software\\Classes\\CLSID\\");
		StringCchCat(classIdKey, maxLen, classId);
	}

	void ProgIdSubKey(LPCWSTR progId, RegisterMode mode, WCHAR* progIdKey, int maxLen)
	{
		StringCchCopy(progIdKey, maxLen, L"Software\\Classes\\");
		StringCchCat(progIdKey, maxLen, progId);
	}

	BOOL SetKeyAndValue(HKEY hKeyRoot, LPCWSTR pszPath, LPCWSTR pszSubkey1, LPCWSTR pszSubkey2, LPCWSTR pszSubkey3, LPCWSTR pszvalueName, LPCWSTR pszValue)
	{
		HKEY hKey;
		WCHAR szKeyBuf[1024];

		StringCchCopy(szKeyBuf, 1024, pszPath);

		if (pszSubkey1 != NULL)
		{
			StringCchCat(szKeyBuf, 1024, L"\\");
			StringCchCat(szKeyBuf, 1024, pszSubkey1);
		}
		if (pszSubkey2 != NULL)
		{
			StringCchCat(szKeyBuf, 1024, L"\\");
			StringCchCat(szKeyBuf, 1024, pszSubkey2);
		}
		if (pszSubkey3 != NULL)
		{
			StringCchCat(szKeyBuf, 1024, L"\\");
			StringCchCat(szKeyBuf, 1024, pszSubkey3);
		}

		// if its return 5 - E_ACCESS_DENIED
		long lResult = RegCreateKeyEx(hKeyRoot, szKeyBuf, 0, NULL, REG_OPTION_NON_VOLATILE, _regKeyOptions, NULL, &hKey, NULL);
		if (lResult != ERROR_SUCCESS)
			return FALSE;

		if (pszValue != NULL)
		{
#if UNICODE
			RegSetValueEx(hKey, pszvalueName, 0, REG_SZ, (BYTE*)pszValue, lstrlen(pszValue) * 2 + 2);
#else
			RegSetValueEx(hKey, pszvalueName, 0, REG_SZ, (BYTE*)pszValue, lstrlen(pszValue) + 1);
#endif
		}

		RegCloseKey(hKey);
		return TRUE;
	}

	LONG RecursiveDeleteKey(HKEY hKeyParent, LPCWSTR pszKeyChild, LPCWSTR pszKeyChild2)
	{
		WCHAR szKeyBuf[1024];
		StringCchCopy(szKeyBuf, 1024, pszKeyChild);
		StringCchCat(szKeyBuf, 1024, L"\\");
		StringCchCat(szKeyBuf, 1024, pszKeyChild2);
		return RecursiveDeleteKey(hKeyParent, szKeyBuf);
	}

	LONG RecursiveDeleteKey(HKEY hKeyParent, LPCWSTR pszKeyChild, LPCWSTR pszKeyChild2, LPCWSTR pszKeyChild3)
	{
		WCHAR szKeyBuf[1024];
		StringCchCopy(szKeyBuf, 1024, pszKeyChild);
		StringCchCat(szKeyBuf, 1024, L"\\");
		StringCchCat(szKeyBuf, 1024, pszKeyChild2);
		StringCchCat(szKeyBuf, 1024, L"\\");
		StringCchCat(szKeyBuf, 1024, pszKeyChild3);
		return RecursiveDeleteKey(hKeyParent, szKeyBuf);
	}

	LONG RecursiveDeleteKey(HKEY hKeyParent, LPCWSTR pszKeyChild)
	{
		HKEY hKeyChild;
		LONG lRes = RegOpenKeyEx(hKeyParent, pszKeyChild, 0, _regKeyOptions, &hKeyChild);
		if (lRes == ERROR_PATH_NOT_FOUND)
			return 0;
		if (lRes != ERROR_SUCCESS)
			return lRes;

		FILETIME time;
		WCHAR szBuffer[256];
		DWORD dwSize = 256;
		while (RegEnumKeyEx(hKeyChild, 0, szBuffer, &dwSize, NULL, NULL, NULL, &time) == S_OK)
		{
			lRes = RecursiveDeleteKey(hKeyChild, szBuffer);
			if (lRes != ERROR_SUCCESS)
			{
				RegCloseKey(hKeyChild);
				return lRes;
			}
			dwSize = 256;
		}

		RegCloseKey(hKeyChild);
		return RegDeleteKey(hKeyParent, pszKeyChild);
	}
}
