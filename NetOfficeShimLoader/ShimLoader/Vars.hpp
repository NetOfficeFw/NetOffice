#ifndef VARS__H
#define VARS__H

#pragma once
//
// Addin Aggregator
//
LPCWSTR TargetManagedAggregator_AssemblyName			  = L"NetOffice, PublicKeyToken=82590859a0ddadaf";
LPCWSTR TargetManagedAggregator_ClassName			  = L"NetOffice.Tools.Isolation.ManagedInnerComAggregator";
LPCWSTR TargetManagedAggregator_AppDomain_FriendlyName = L"";
LPCWSTR TargetManagedAggregator_AppDomain_BaseFolder	  = L"";


//
// Addin Target
//
LPCWSTR Target_AssemblyName							  = L"InnerAddin, PublicKeyToken=6153aeeaee4248b8";
LPCWSTR Target_AssemblyFileName						  = L"InnerAddin.dll";
LPCWSTR Target_ConnectClassName						  = L"InnerAddin.Connect";
LPCWSTR Target_ConfigFileName						  = L"InnerAddin.dll.config";


//
// Register Values
//
LPCWSTR  ShimProxy_CLSID						 = L"{FF724928-8E6B-4A1E-97F3-C6B9A944154C}";
LPCWSTR  ShimProxy_ProgID							 = L"ZLoaderShim.Connect";
LPCWSTR  ShimProxy_Version							 = L"1.0.0.0";
LPCWSTR  ShimProxy_FriendlyName						 = L"NetOffice Generic COM Shim";
LPCWSTR  ShimProxy_Description						 = L"NetOffice Generic COM Shim";
DWORD    ShimProxy_LoadBehavior						 = 3;
DWORD    ShimProxy_CommandLineSafe					 = 0;
LPCWSTR* ShimProxy_Host_Application					 = NULL;


//
// Update Aggregator
//
LPCWSTR UpdateManagedAggregator_AssemblyName			 = L"InnerUpdate, PublicKeyToken=e58b77e9e2189611";
LPCWSTR UpdateManagedAggregator_ClassName			 = L"NetOffice.Tools.Isolation.ManagedInnerUpdateAggregator";


//
// Update Handler
//
LPCWSTR Update_AssemblyName						    = L"InnerUpdate, PublicKeyToken=e58b77e9e2189611";
LPCWSTR Update_AssemblyFileName					    = L"InnerUpdate.dll";
LPCWSTR Update_ConnectClassName					    = L"InnerUpdate.Connect";
LPCWSTR Update_ConfigFileName					    = L"InnerUpdate.dll.config";


//
// Settings
//
BOOL ENABLE_SHIM										= TRUE;
BOOL ENABLE_SELF_REGISTRATION								= TRUE;
BOOL ENABLE_TARGET_REGISTRATION						= TRUE;
int	 SELF_REGISTER_MODE								= 0; // System = 0, User = 1, SystemComponentAndUserAddin = 2
BOOL ENABLE_BLIND_AGGREGATION						= FALSE;
BOOL ENABLE_OUTER_UPDATE_AGGREGATOR					= TRUE;
BOOL ENABLE_DEBUG_MESSAGE_BOX						= TRUE;


//
// Defaults
//
HRESULT EXTENSIBILITY_DEFAULT_RESULT					= S_OK;
HRESULT EXTENSIBILITY_FAIL_RESULT					= S_OK;

#endif !VARS__H
