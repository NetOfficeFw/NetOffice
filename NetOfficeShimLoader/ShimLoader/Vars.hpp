#pragma once
#include "stdafx.h"

//
// Register Values
//
static const GUID ShimProxy_CLSID = { 0xff724928, 0x8e6b, 0x4a1e, 0x97, 0xf3, 0xc6, 0xb9, 0xa9, 0x44, 0x15, 0x4c };
static const LPCWSTR ShimProxy_CLSID_Text = L"{FF724928-8E6B-4A1E-97F3-C6B9A944154C}";
static const LPCWSTR ShimProxy_ProgID = L"ZLoaderShim.Connect";
static const LPCWSTR ShimProxy_Version = L"1.0.0.0";
static const LPCWSTR ShimProxy_FriendlyName = L"NetOffice Generic COM Shim";
static const LPCWSTR ShimProxy_Description = L"NetOffice Generic COM Shim";
static const DWORD ShimProxy_LoadBehavior = 3;
static const DWORD ShimProxy_CommandLineSafe = 0;
static LPCWSTR* ShimProxy_Host_Application = NULL;

//
// Addin Target
//
static const LPCWSTR Target_AssemblyName = L"InnerAddin, PublicKeyToken=6153aeeaee4248b8";
static const LPCWSTR Target_AssemblyFileName = L"InnerAddin.dll";
static const LPCWSTR Target_ConnectClassName = L"InnerAddin.Addin";
static const LPCWSTR Target_ConfigFileName = L"InnerAddin.dll.config";

//
// Aggregator
//
static const LPCWSTR ManagedAggregator_AssemblyName = L"NetOffice, PublicKeyToken=82590859a0ddadaf";
static const LPCWSTR ManagedAggregator_ClassName = L"NetOffice.Tools.Isolation.ManagedInnerAggregator";

//
// Settings
//
static const BOOL ENABLE_SHIM = TRUE;
static const BOOL ENABLE_SELF_REGISTRATION = TRUE;
static const int SELF_REGISTER_MODE = 2; // System = 0, User = 1, SystemComponentAndUserAddin = 2
static const BOOL ENABLE_BLIND_AGGREGATION = FALSE;
static const BOOL ENABLE_OUTER_UPDATE_AGGREGATOR = TRUE;
static const BOOL ENABLE_DEBUG_MESSAGE_BOX = TRUE;

//
// Defaults
//
static const HRESULT EXTENSIBILITY_DEFAULT_RESULT = S_OK;
static const HRESULT EXTENSIBILITY_FAIL_RESULT = S_OK;
