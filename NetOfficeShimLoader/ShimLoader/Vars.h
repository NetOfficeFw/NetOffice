#pragma once
#include "CustomRegisterValue.h"

using namespace NetOffice_ShimLoader_Register;

extern WCHAR* LogFile_Path;
extern WCHAR* LogFile_Register_Path;
extern WCHAR* LogFile_UnRegister_Path;

extern BOOL Internal_LogError_MessageBoxes_Enabled;

//
// Managed Addin Aggregator
//
extern WCHAR* TargetManagedAggregator_Folder;
extern WCHAR* TargetManagedAggregator_AssemblyName;
extern WCHAR* TargetManagedAggregator_ClassName;
extern WCHAR* TargetManagedAggregator_AppDomain_FriendlyName;
extern WCHAR* TargetManagedAggregator_AppDomain_BaseFolder;


//
// Managed Addin Target
//
extern WCHAR* Target_AssemblyName;
extern WCHAR* Target_AssemblyFileName;
extern WCHAR* Target_ConnectClassName;
extern WCHAR* Target_ConfigFileName;


//
// Register Values
//
extern WCHAR*  ShimProxy_CLSID;
extern WCHAR*  ShimProxy_ProgID;
extern WCHAR*  ShimProxy_Version;
extern WCHAR*  ShimProxy_FriendlyName;
extern WCHAR*  ShimProxy_Description;
extern DWORD   ShimProxy_LoadBehavior;
extern DWORD   ShimProxy_CommandLineSafe;
extern LPCWSTR*  ShimProxy_Host_Application;
extern size_t  ShimProxy_Host_Application_Length;
extern PCustomRegisterValue* Custom_Register_Values;
extern size_t  Custom_Register_Values_Length;

//
// Managed Update Aggregator
//
extern WCHAR* UpdateManagedAggregator_Folder;
extern WCHAR* UpdateManagedAggregator_AssemblyName;
extern WCHAR* UpdateManagedAggregator_ClassName;
extern WCHAR* UpdateManagedAggregator_AppDomain_FriendlyName;
extern WCHAR* UpdateManagedAggregator_AppDomain_BaseFolder;


//
// Managed Update Handler
//
extern WCHAR* Update_AssemblyName;
extern WCHAR* Update_AssemblyFileName;
extern WCHAR* Update_ConnectClassName;
extern WCHAR* Update_ConfigFileName;


//
// Settings
//
extern BOOL ENABLE_SHIM;
extern BOOL ENABLE_SELF_REGISTRATION;
extern BOOL ENABLE_TARGET_REGISTRATION;
extern BOOL ENABLE_ADDIN_REGISTRATION;
extern int	SELF_REGISTER_MODE;
extern BOOL ENABLE_BLIND_AGGREGATION;
extern BOOL ENABLE_OUTER_UPDATE_AGGREGATOR;
extern BOOL ENABLE_DEBUG_MESSAGE_BOX;


//
// Defaults
//
extern HRESULT EXTENSIBILITY_DEFAULT_RESULT;
extern HRESULT EXTENSIBILITY_FAIL_RESULT;
