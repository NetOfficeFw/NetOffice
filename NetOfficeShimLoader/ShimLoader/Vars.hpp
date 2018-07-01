#pragma once
#include "stdafx.h"

// Register Values
//
static const GUID ShimProxy_CLSID = { 0xff724928, 0x8e6b, 0x4a1e, 0x97, 0xf3, 0xc6, 0xb9, 0xa9, 0x44, 0x15, 0x4c };
static const LPCWSTR ShimProxy_CLSID_Text = L"{FF724928-8E6B-4A1E-97F3-C6B9A944154C}";
static const LPCWSTR ShimProxy_ProgID = L"LoaderShim.Connect";
static const LPCWSTR ShimProxy_Description = L"NetOffice Generic COM Shim";
static const LPCWSTR ShimProxy_Host_Application = L"PowerPoint";

//
// Addin Target
//
static LPCWSTR Target_AssemblyName = L"InnerAddin, PublicKeyToken=6153aeeaee4248b8";
static LPCWSTR Target_AssemblyFileName = L"InnerAddin.dll";
static LPCWSTR Target_ConnectClassName = L"InnerAddin.Addin";
static LPCWSTR Target_ConfigFileName = L"InnerAddin.dll.config";

//
// Aggregator
//
static const LPCWSTR ManagedAggregator_AssemblyName = L"NetOffice, PublicKeyToken=82590859a0ddadaf";
static const LPCWSTR ManagedAggregator_ClassName = L"NetOffice.Tools.Isolation.ManagedInnerAggregator";

//
// Settings
//
static bool ENABLE_SELF_REGISTRATION = true;
static bool ENABLE_BLIND_AGGREGATION = false;
static bool ENABLE_OUTER_UPDATE_AGGREGATOR = true;