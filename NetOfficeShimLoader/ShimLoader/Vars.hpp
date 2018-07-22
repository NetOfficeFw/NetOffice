#pragma once
#include "CustomRegisterValue.h"

using namespace NetOffice_ShimLoader_Register;

WCHAR* LogFile_Path;
WCHAR* LogFile_Register_Path;
WCHAR* LogFile_UnRegister_Path;

BOOL Internal_LogError_MessageBoxes_Enabled;

//
// Managed Addin Aggregator
//
WCHAR* TargetManagedAggregator_Folder						= nullptr;
WCHAR* TargetManagedAggregator_AssemblyName					= nullptr; // = L"NetOffice, PublicKeyToken=82590859a0ddadaf";
WCHAR* TargetManagedAggregator_ClassName					= nullptr; // = L"NetOffice.Tools.Isolation.ManagedInnerComAggregator";
WCHAR* TargetManagedAggregator_AppDomain_FriendlyName		= nullptr;
WCHAR* TargetManagedAggregator_AppDomain_BaseFolder			= nullptr;


//
// Managed Addin Target
//
WCHAR* Target_AssemblyName									= nullptr; // = L"InnerAddin, PublicKeyToken=6153aeeaee4248b8";
WCHAR* Target_AssemblyFileName								= nullptr; // = L"InnerAddin.dll";
WCHAR* Target_ConnectClassName								= nullptr; // = L"InnerAddin.Connect";
WCHAR* Target_ConfigFileName								= nullptr; // = L"InnerAddin.dll.config";


//
// Register Values
//
WCHAR*  ShimProxy_CLSID										= nullptr; // = L"{FF724928-8E6B-4A1E-97F3-C6B9A944154C}";
WCHAR*  ShimProxy_ProgID									= nullptr; // = L"ZLoaderShim.Connect";
WCHAR*  ShimProxy_Version									= nullptr; // = L"1.0.0.0";
WCHAR*  ShimProxy_FriendlyName								= nullptr; // = L"NetOffice Generic COM Shim";
WCHAR*  ShimProxy_Description								= nullptr; // = L"NetOffice Generic COM Shim";
DWORD   ShimProxy_LoadBehavior								= 3;
DWORD   ShimProxy_CommandLineSafe							= 0;
LPCWSTR*  ShimProxy_Host_Application						= nullptr;
size_t  ShimProxy_Host_Application_Length					= 0;
PCustomRegisterValue* Custom_Register_Values				= nullptr;
size_t  Custom_Register_Values_Length						= 0;

//
// Managed Update Aggregator
//
WCHAR* UpdateManagedAggregator_Folder						= nullptr;
WCHAR* UpdateManagedAggregator_AssemblyName					= L"InnerUpdate, PublicKeyToken=e58b77e9e2189611";
WCHAR* UpdateManagedAggregator_ClassName					= L"NetOffice.Tools.Isolation.ManagedInnerUpdateAggregator";
WCHAR* UpdateManagedAggregator_AppDomain_FriendlyName		= nullptr;
WCHAR* UpdateManagedAggregator_AppDomain_BaseFolder			= nullptr;


//
// Managed Update Handler
//
WCHAR* Update_AssemblyName									= L"InnerUpdate, PublicKeyToken=e58b77e9e2189611";
WCHAR* Update_AssemblyFileName								= L"InnerUpdate.dll";
WCHAR* Update_ConnectClassName								= L"InnerUpdate.Connect";
WCHAR* Update_ConfigFileName								= L"InnerUpdate.dll.config";


//
// Settings
//
BOOL ENABLE_SHIM											= TRUE;
BOOL ENABLE_SELF_REGISTRATION								= TRUE;
BOOL ENABLE_TARGET_REGISTRATION								= TRUE;
BOOL ENABLE_ADDIN_REGISTRATION								= TRUE;
int	 SELF_REGISTER_MODE										= 0; // System = 0, User = 1, SystemComponentAndUserAddin = 2
BOOL ENABLE_BLIND_AGGREGATION								= FALSE;
BOOL ENABLE_OUTER_UPDATE_AGGREGATOR							= TRUE;
BOOL ENABLE_DEBUG_MESSAGE_BOX								= TRUE;


//
// Defaults
//
HRESULT EXTENSIBILITY_DEFAULT_RESULT						= S_OK;
HRESULT EXTENSIBILITY_FAIL_RESULT							= S_OK;
