#pragma once
#include "stdafx.h"

__interface __declspec(uuid("000c033b-0000-0000-c000-000000000046"))
	_CustomTaskPane : IDispatch
{
	__declspec(property(get = GetTitle)) _bstr_t Title;
	__declspec(property(get = GetApplication)) IDispatchPtr Application;
	__declspec(property(get = GetWindow)) IDispatchPtr Window;
	__declspec(property(get = GetVisible, put = PutVisible)) VARIANT_BOOL Visible;
	__declspec(property(get = GetContentControl)) IDispatchPtr ContentControl;
	__declspec(property(get = GetHeight, put = PutHeight)) int Height;
	__declspec(property(get = GetWidth, put = PutWidth)) int Width;
	__declspec(property(get = GetDockPosition, put = PutDockPosition)) enum MsoCTPDockPosition DockPosition;
	__declspec(property(get = GetDockPositionRestrict, put = PutDockPositionRestrict)) enum MsoCTPDockPositionRestrict DockPositionRestrict;

	virtual HRESULT __stdcall get_Title(
		/*[out,retval]*/ BSTR * prop) = 0;
	virtual HRESULT __stdcall get_Application(
		/*[out,retval]*/ IDispatch * * prop) = 0;
	virtual HRESULT __stdcall get_Window(
		/*[out,retval]*/ IDispatch * * prop) = 0;
	virtual HRESULT __stdcall get_Visible(
		/*[out,retval]*/ VARIANT_BOOL * prop) = 0;
	virtual HRESULT __stdcall put_Visible(
		/*[in]*/ VARIANT_BOOL prop) = 0;
	virtual HRESULT __stdcall get_ContentControl(
		/*[out,retval]*/ IDispatch * * prop) = 0;
	virtual HRESULT __stdcall get_Height(
		/*[out,retval]*/ int * prop) = 0;
	virtual HRESULT __stdcall put_Height(
		/*[in]*/ int prop) = 0;
	virtual HRESULT __stdcall get_Width(
		/*[out,retval]*/ int * prop) = 0;
	virtual HRESULT __stdcall put_Width(
		/*[in]*/ int prop) = 0;
	virtual HRESULT __stdcall get_DockPosition(
		/*[out,retval]*/ enum MsoCTPDockPosition * prop) = 0;
	virtual HRESULT __stdcall put_DockPosition(
		/*[in]*/ enum MsoCTPDockPosition prop) = 0;
	virtual HRESULT __stdcall get_DockPositionRestrict(
		/*[out,retval]*/ enum MsoCTPDockPositionRestrict * prop) = 0;
	virtual HRESULT __stdcall put_DockPositionRestrict(
		/*[in]*/ enum MsoCTPDockPositionRestrict prop) = 0;
	virtual HRESULT __stdcall Delete() = 0;
};
// 000c033b-0000-0000-c000-000000000046
static const GUID IID_CustomTaskPane = __uuidof(_CustomTaskPane);

__interface __declspec(uuid("000c033d-0000-0000-c000-000000000046"))
	ICTPFactory : IDispatch
{
	STDMETHOD(CreateCTP)
		(THIS_/*[in]*/ BSTR CTPAxID,
			/*[in]*/ BSTR CTPTitle,
			/*[in]*/ VARIANT CTPParentWindow,
			/*[out,retval]*/ struct _CustomTaskPane * * CTPInst) PURE;
};
// 000c033d-0000-0000-c000-000000000046
static const GUID IID_ICTPFactory = __uuidof(ICTPFactory);

__interface __declspec(uuid("000c033e-0000-0000-c000-000000000046"))
	ICustomTaskPaneConsumer : IDispatch
{
	STDMETHOD(CTPFactoryAvailable) (THIS_/*[in]*/ ICTPFactory* CTPFactoryInst) PURE;
};
// 000c033e-0000-0000-c000-000000000046
static const GUID IID_ICustomTaskPaneConsumer = __uuidof(ICustomTaskPaneConsumer);
