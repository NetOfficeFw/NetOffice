#pragma once
#include "stdafx.h"

//
// Addin Aggregator
//
static LPCWSTR TargetManagedAggregator_AssemblyName			  = L"NetOffice, PublicKeyToken=82590859a0ddadaf";
static LPCWSTR TargetManagedAggregator_ClassName			  = L"NetOffice.Tools.Isolation.ManagedInnerComAggregator";
static LPCWSTR TargetManagedAggregator_AppDomain_FriendlyName = L"";
static LPCWSTR TargetManagedAggregator_AppDomain_BaseFolder	  = L"";


//
// Addin Target
//
static LPCWSTR Target_AssemblyName							  = L"InnerAddin, PublicKeyToken=6153aeeaee4248b8";
static LPCWSTR Target_AssemblyFileName						  = L"InnerAddin.dll";
static LPCWSTR Target_ConnectClassName						  = L"InnerAddin.Connect";
static LPCWSTR Target_ConfigFileName						  = L"InnerAddin.dll.config";


//
// Register Values
//
static GUID	    ShimProxy_CLSID = { 0xff724928, 0x8e6b, 0x4a1e, 0x97, 0xf3, 0xc6, 0xb9, 0xa9, 0x44, 0x15, 0x4c };
static LPCWSTR  ShimProxy_CLSID_Text						 = L"{FF724928-8E6B-4A1E-97F3-C6B9A944154C}";
static LPCWSTR  ShimProxy_ProgID							 = L"ZLoaderShim.Connect";
static LPCWSTR  ShimProxy_Version							 = L"1.0.0.0";
static LPCWSTR  ShimProxy_FriendlyName						 = L"NetOffice Generic COM Shim";
static LPCWSTR  ShimProxy_Description						 = L"NetOffice Generic COM Shim";
static DWORD    ShimProxy_LoadBehavior						 = 3;
static DWORD    ShimProxy_CommandLineSafe					 = 0;
static LPCWSTR* ShimProxy_Host_Application					 = NULL;


//
// Update Aggregator
//
static LPCWSTR UpdateManagedAggregator_AssemblyName			 = L"InnerUpdate, PublicKeyToken=e58b77e9e2189611";
static LPCWSTR UpdateManagedAggregator_ClassName			 = L"NetOffice.Tools.Isolation.ManagedInnerUpdateAggregator";


//
// Update Handler
//
static LPCWSTR Update_AssemblyName						    = L"InnerUpdate, PublicKeyToken=e58b77e9e2189611";
static LPCWSTR Update_AssemblyFileName					    = L"InnerUpdate.dll";
static LPCWSTR Update_ConnectClassName					    = L"InnerUpdate.Connect";
static LPCWSTR Update_ConfigFileName					    = L"InnerUpdate.dll.config";


//
// Settings
//
const BOOL ENABLE_SHIM									    = TRUE;
const BOOL ENABLE_SELF_REGISTRATION						    = TRUE;
const int  SELF_REGISTER_MODE							    = 0; // System = 0, User = 1, SystemComponentAndUserAddin = 2
const BOOL ENABLE_BLIND_AGGREGATION						    = FALSE;
const BOOL ENABLE_OUTER_UPDATE_AGGREGATOR				    = TRUE;
const BOOL ENABLE_DEBUG_MESSAGE_BOX						    = TRUE;


//
// Defaults
//
const HRESULT EXTENSIBILITY_DEFAULT_RESULT				   = S_OK;
const HRESULT EXTENSIBILITY_FAIL_RESULT					   = S_OK;
