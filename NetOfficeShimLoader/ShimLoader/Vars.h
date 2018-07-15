#pragma once
#include "CustomRegisterValue.h"

//
// Addin Aggregator
//
extern LPCWSTR TargetManagedAggregator_AssemblyName;
extern LPCWSTR TargetManagedAggregator_ClassName;
extern LPCWSTR TargetManagedAggregator_AppDomain_FriendlyName;
extern LPCWSTR TargetManagedAggregator_AppDomain_BaseFolder;


//
// Addin Target
//
extern LPCWSTR Target_AssemblyName;
extern LPCWSTR Target_AssemblyFileName;
extern LPCWSTR Target_ConnectClassName;
extern LPCWSTR Target_ConfigFileName;


//
// Register Values
//
extern LPCWSTR  ShimProxy_CLSID;
extern LPCWSTR  ShimProxy_ProgID;
extern LPCWSTR  ShimProxy_Version;
extern LPCWSTR  ShimProxy_FriendlyName;
extern LPCWSTR  ShimProxy_Description;
extern DWORD    ShimProxy_LoadBehavior;
extern DWORD    ShimProxy_CommandLineSafe;
extern LPCWSTR* ShimProxy_Host_Application;
extern NetOffice_ShimLoader_Register::PCustomRegisterValue* Custom_Register_Values;

//
// Update Aggregator
//
extern LPCWSTR UpdateManagedAggregator_AssemblyName;
extern LPCWSTR UpdateManagedAggregator_ClassName;


//
// Update Handler
//
extern LPCWSTR Update_AssemblyName;
extern LPCWSTR Update_AssemblyFileName;
extern LPCWSTR Update_ConnectClassName;
extern LPCWSTR Update_ConfigFileName;


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
