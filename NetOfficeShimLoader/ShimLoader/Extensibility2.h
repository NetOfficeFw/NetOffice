#pragma once
#include "stdafx.h"

//
// Definition of IDTExtensibility2
//

enum ext_ConnectMode {
	ext_cm_AfterStartup = 0,
	ext_cm_Startup = 1,
	ext_cm_External = 2,
	ext_cm_CommandLine = 3
};

enum ext_DisconnectMode {
	ext_dm_HostShutdown = 0,
	ext_dm_UserClosed = 1
};

__interface __declspec(uuid("B65AD801-ABAF-11D0-BB8B-00A0C90F2744"))
	IDTExtensibility2 : public IDispatch
{
	STDMETHOD(OnConnection)(THIS_ IDispatch FAR* Application, ext_ConnectMode ConnectMode, IDispatch FAR* AddInInst, LPSAFEARRAY FAR* custom) PURE;
	STDMETHOD(OnDisconnection)(THIS_ ext_DisconnectMode RemoveMode, LPSAFEARRAY FAR* custom) PURE;
	STDMETHOD(OnAddInsUpdate)(THIS_ LPSAFEARRAY FAR* custom) PURE;
	STDMETHOD(OnStartupComplete)(THIS_ LPSAFEARRAY FAR* custom) PURE;
	STDMETHOD(OnBeginShutdown)(THIS_ LPSAFEARRAY FAR* custom) PURE;
};
// B65AD801-ABAF-11D0-BB8B-00A0C90F2744
static const GUID IID_IDTExtensibility2 = { 0xB65AD801L,0xABAF,0x11D0,{ 0xBB,0x8B,0x00,0xA0,0xC9,0x0F,0x27,0x44 } };