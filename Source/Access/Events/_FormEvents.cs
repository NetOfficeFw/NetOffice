using System;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
using NetOffice;
using NetOffice.Attributes;

namespace NetOffice.AccessApi.Events
{	
	#pragma warning disable
	
	#region SinkPoint Interface

	[SupportByVersion("Access", 9,10,11,12,14,15,16)]
    [InternalEntity(InternalEntityKind.ComEventInterface)]
    [ComImport, Guid("331FDCFB-CF31-11CD-8701-00AA003F0F07"), InterfaceType(ComInterfaceType.InterfaceIsIDispatch), TypeLibType((short)0x1010)]
	public interface _FormEvents
	{
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(2067)]
		void Load();

		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(2058)]
		void Current();

		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(2059)]
		void BeforeInsert([In] [Out] ref object cancel);

		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(2060)]
		void AfterInsert();

		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(2061)]
		void BeforeUpdate([In] [Out] ref object cancel);

		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(2062)]
		void AfterUpdate();

		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
        [SinkArgument("cancel", SinkArgumentType.Int16)]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(2063)]
		void Delete([In] [Out] ref object cancel);

		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
        [SinkArgument("cancel", SinkArgumentType.Int16)]
        [SinkArgument("response", SinkArgumentType.Int16)]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(2064)]
		void BeforeDelConfirm([In] [Out] ref object cancel, [In] [Out] ref object response);

		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
        [SinkArgument("status", SinkArgumentType.Int16)]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(2065)]
		void AfterDelConfirm([In] [Out] ref object status);

		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
        [SinkArgument("cancel", SinkArgumentType.Int16)]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(2066)]
		void Open([In] [Out] ref object cancel);

		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(2068)]
		void Resize();

		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
        [SinkArgument("cancel", SinkArgumentType.Int16)]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(2069)]
		void Unload([In] [Out] ref object cancel);

		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(2070)]
		void Close();

		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(2071)]
		void Activate();

		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(2072)]
		void Deactivate();

		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(2073)]
		void GotFocus();

		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(2074)]
		void LostFocus();

		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(-600)]
		void Click();

		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
        [SinkArgument("cancel", SinkArgumentType.Int16)]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(-601)]
		void DblClick([In] [Out] ref object cancel);

		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
        [SinkArgument("button", SinkArgumentType.Int16)]
        [SinkArgument("shift", SinkArgumentType.Int16)]
        [SinkArgument("x", SinkArgumentType.Single)]
        [SinkArgument("y", SinkArgumentType.Single)]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(-605)]
		void MouseDown([In] [Out] ref object button, [In] [Out] ref object shift, [In] [Out] ref object x, [In] [Out] ref object y);

		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
        [SinkArgument("button", SinkArgumentType.Int16)]
        [SinkArgument("shift", SinkArgumentType.Int16)]
        [SinkArgument("x", SinkArgumentType.Single)]
        [SinkArgument("y", SinkArgumentType.Single)]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(-606)]
		void MouseMove([In] [Out] ref object button, [In] [Out] ref object shift, [In] [Out] ref object x, [In] [Out] ref object y);

		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
        [SinkArgument("button", SinkArgumentType.Int16)]
        [SinkArgument("shift", SinkArgumentType.Int16)]
        [SinkArgument("x", SinkArgumentType.Single)]
        [SinkArgument("y", SinkArgumentType.Single)]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(-607)]
		void MouseUp([In] [Out] ref object button, [In] [Out] ref object shift, [In] [Out] ref object x, [In] [Out] ref object y);

		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
        [SinkArgument("keyCode", SinkArgumentType.Int16)]
        [SinkArgument("shift", SinkArgumentType.Int16)]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(-602)]
		void KeyDown([In] [Out] ref object keyCode, [In] [Out] ref object shift);

		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
        [SinkArgument("keyAscii", SinkArgumentType.Int16)]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(-603)]
		void KeyPress([In] [Out] ref object keyAscii);

		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
        [SinkArgument("keyCode", SinkArgumentType.Int16)]
        [SinkArgument("shift", SinkArgumentType.Int16)]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(-604)]
		void KeyUp([In] [Out] ref object keyCode, [In] [Out] ref object shift);

		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
        [SinkArgument("dataErr", SinkArgumentType.Int16)]
        [SinkArgument("response", SinkArgumentType.Int16)]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(2083)]
		void Error([In] [Out] ref object dataErr, [In] [Out] ref object response);

		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(2084)]
		void Timer();

		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
        [SinkArgument("cancel", SinkArgumentType.Int16)]
        [SinkArgument("filterType", SinkArgumentType.Int16)]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(2155)]
		void Filter([In] [Out] ref object cancel, [In] [Out] ref object filterType);

		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
        [SinkArgument("cancel", SinkArgumentType.Int16)]
        [SinkArgument("filterType", SinkArgumentType.Int16)]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(2156)]
		void ApplyFilter([In] [Out] ref object cancel, [In] [Out] ref object applyType);

		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
        [SinkArgument("cancel", SinkArgumentType.Int16)]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(2205)]
		void Dirty([In] [Out] ref object cancel);

		[SupportByVersion("Access", 10,11,12,14,15,16)]
        [SinkArgument("cancel", SinkArgumentType.Int16)]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(2145)]
		void Undo([In] [Out] ref object cancel);

		[SupportByVersion("Access", 10,11,12,14,15,16)]
        [SinkArgument("cancel", SinkArgumentType.Int16)]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(2334)]
		void RecordExit([In] [Out] ref object cancel);

		[SupportByVersion("Access", 10,11,12,14,15,16)]
        [SinkArgument("cancel", SinkArgumentType.Int16)]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(2369)]
		void BeginBatchEdit([In] [Out] ref object cancel);

		[SupportByVersion("Access", 10,11,12,14,15,16)]
        [SinkArgument("cancel", SinkArgumentType.Int16)]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(2370)]
		void UndoBatchEdit([In] [Out] ref object cancel);

        [SupportByVersion("Access", 10,11,12,14,15,16)]
        [SinkArgument("cancel", SinkArgumentType.Int16)]
        [SinkArgument("connection", typeof(NetOffice.ADODBApi.Connection))]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(2371)]
		void BeforeBeginTransaction([In] [Out] ref object cancel, [In] [Out, MarshalAs(UnmanagedType.IDispatch)] object connection);

		[SupportByVersion("Access", 10,11,12,14,15,16)]
        [SinkArgument("connection", typeof(NetOffice.ADODBApi.Connection))]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(2372)]
		void AfterBeginTransaction([In] [Out, MarshalAs(UnmanagedType.IDispatch)] object connection);

		[SupportByVersion("Access", 10,11,12,14,15,16)]
        [SinkArgument("cancel", SinkArgumentType.Int16)]
        [SinkArgument("connection", typeof(NetOffice.ADODBApi.Connection))]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(2373)]
		void BeforeCommitTransaction([In] [Out] ref object cancel, [In] [Out, MarshalAs(UnmanagedType.IDispatch)] object connection);

		[SupportByVersion("Access", 10,11,12,14,15,16)]
        [SinkArgument("connection", typeof(NetOffice.ADODBApi.Connection))]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(2374)]
		void AfterCommitTransaction([In] [Out, MarshalAs(UnmanagedType.IDispatch)] object connection);

		[SupportByVersion("Access", 10,11,12,14,15,16)]
        [SinkArgument("connection", typeof(NetOffice.ADODBApi.Connection))]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(2375)]
		void RollbackTransaction([In] [Out, MarshalAs(UnmanagedType.IDispatch)] object connection);

		[SupportByVersion("Access", 10,11,12,14,15,16)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(2383)]
		void OnConnect();

		[SupportByVersion("Access", 10,11,12,14,15,16)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(2384)]
		void OnDisconnect();

		[SupportByVersion("Access", 10,11,12,14,15,16)]
        [SinkArgument("reason", SinkArgumentType.Int32)]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(2385)]
		void PivotTableChange([In] object reason);

		[SupportByVersion("Access", 10,11,12,14,15,16)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(2386)]
		void Query();

		[SupportByVersion("Access", 10,11,12,14,15,16)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(2387)]
		void BeforeQuery();

		[SupportByVersion("Access", 10,11,12,14,15,16)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(2388)]
		void SelectionChange();

		[SupportByVersion("Access", 10,11,12,14,15,16)]
        [SinkArgument("cancel", SinkArgumentType.UnknownProxy)]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(2389)]
		void CommandBeforeExecute([In] object command, [In, MarshalAs(UnmanagedType.IDispatch)] object cancel);

		[SupportByVersion("Access", 10,11,12,14,15,16)]
        [SinkArgument("_checked", SinkArgumentType.UnknownProxy)]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(2390)]
		void CommandChecked([In] object command, [In, MarshalAs(UnmanagedType.IDispatch)] object _checked);

		[SupportByVersion("Access", 10,11,12,14,15,16)]
        [SinkArgument("enabled", SinkArgumentType.UnknownProxy)]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(2391)]
		void CommandEnabled([In] object command, [In, MarshalAs(UnmanagedType.IDispatch)] object enabled);

		[SupportByVersion("Access", 10,11,12,14,15,16)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(2392)]
		void CommandExecute([In] object command);

		[SupportByVersion("Access", 10,11,12,14,15,16)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(2394)]
		void DataSetChange();

		[SupportByVersion("Access", 10,11,12,14,15,16)]
        [SinkArgument("screenTipText", SinkArgumentType.UnknownProxy)]
        [SinkArgument("sourceObject", SinkArgumentType.UnknownProxy)]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(2395)]
		void BeforeScreenTip([In, MarshalAs(UnmanagedType.IDispatch)] object screenTipText, [In, MarshalAs(UnmanagedType.IDispatch)] object sourceObject);

		[SupportByVersion("Access", 10,11,12,14,15,16)]
        [SinkArgument("drawObject", SinkArgumentType.UnknownProxy)]
        [SinkArgument("chartObject", SinkArgumentType.UnknownProxy)]
        [SinkArgument("cancel", SinkArgumentType.UnknownProxy)]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(2399)]
		void BeforeRender([In, MarshalAs(UnmanagedType.IDispatch)] object drawObject, [In, MarshalAs(UnmanagedType.IDispatch)] object chartObject, [In, MarshalAs(UnmanagedType.IDispatch)] object cancel);

		[SupportByVersion("Access", 10,11,12,14,15,16)]
        [SinkArgument("drawObject", SinkArgumentType.UnknownProxy)]
        [SinkArgument("chartObject", SinkArgumentType.UnknownProxy)]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(2397)]
		void AfterRender([In, MarshalAs(UnmanagedType.IDispatch)] object drawObject, [In, MarshalAs(UnmanagedType.IDispatch)] object chartObject);

		[SupportByVersion("Access", 10,11,12,14,15,16)]
        [SinkArgument("drawObject", SinkArgumentType.UnknownProxy)]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(2396)]
		void AfterFinalRender([In, MarshalAs(UnmanagedType.IDispatch)] object drawObject);

		[SupportByVersion("Access", 10,11,12,14,15,16)]
        [SinkArgument("drawObject", SinkArgumentType.UnknownProxy)]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(2398)]
		void AfterLayout([In, MarshalAs(UnmanagedType.IDispatch)] object drawObject);

		[SupportByVersion("Access", 10,11,12,14,15,16)]
        [SinkArgument("drawObject", SinkArgumentType.Bool)]
        [SinkArgument("count", SinkArgumentType.Int32)]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(2401)]
		void MouseWheel([In] object page, [In] object count);

		[SupportByVersion("Access", 10,11,12,14,15,16)]
        [SinkArgument("reason", SinkArgumentType.Int32)]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(2402)]
		void ViewChange([In] object reason);

		[SupportByVersion("Access", 10,11,12,14,15,16)]
        [SinkArgument("reason", SinkArgumentType.Int32)]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(2403)]
		void DataChange([In] object reason);
	}

    #endregion

    #region SinkHelper

    [InternalEntity(InternalEntityKind.SinkHelper)]
    [ComVisible(true), ClassInterface(ClassInterfaceType.None), TypeLibType(TypeLibTypeFlags.FHidden)]
	public class _FormEvents_SinkHelper : SinkHelper, _FormEvents
	{
		#region Static
		
		public static readonly string Id = "331FDCFB-CF31-11CD-8701-00AA003F0F07";
		
		#endregion
				
		#region Ctor

		public _FormEvents_SinkHelper(ICOMObject eventClass, IConnectionPoint connectPoint): base(eventClass)
		{
			SetupEventBinding(connectPoint);
		}
		
		#endregion

		#region _FormEvents
		
		public void Load()
        {
            if (!Validate("Load"))
            {
                return;
            }

			object[] paramsArray = new object[0];
			EventBinding.RaiseCustomEvent("Load", ref paramsArray);
		}

		public void Current()
        {
            if (!Validate("Current"))
            {
                return;
            }

			object[] paramsArray = new object[0];
			EventBinding.RaiseCustomEvent("Current", ref paramsArray);
		}

		public void BeforeInsert([In] [Out] ref object cancel)
        {
            if (!Validate("BeforeInsert"))
            {
                Invoker.ReleaseParamsArray(cancel);
                return;
            }

			object[] paramsArray = new object[1];
			paramsArray.SetValue(cancel, 0);
			EventBinding.RaiseCustomEvent("BeforeInsert", ref paramsArray);

			cancel = ToInt16(paramsArray[0]);
		}

		public void AfterInsert()
        {
            if (!Validate("AfterInsert"))
            {
                return;
            }

			object[] paramsArray = new object[0];
			EventBinding.RaiseCustomEvent("AfterInsert", ref paramsArray);
		}

		public void BeforeUpdate([In] [Out] ref object cancel)
        {
            if (!Validate("BeforeUpdate"))
            {
                Invoker.ReleaseParamsArray(cancel);
                return;
            }

			object[] paramsArray = new object[1];
			paramsArray.SetValue(cancel, 0);
			EventBinding.RaiseCustomEvent("BeforeUpdate", ref paramsArray);

			cancel = ToInt16(paramsArray[0]);
		}

		public void AfterUpdate()
        {
            if (!Validate("AfterUpdate"))
            {
                return;
            }
			object[] paramsArray = new object[0];
			EventBinding.RaiseCustomEvent("AfterUpdate", ref paramsArray);
		}

		public void Delete([In] [Out] ref object cancel)
        {
            if (!Validate("Delete"))
            {
                Invoker.ReleaseParamsArray(cancel);
                return;
            }

			object[] paramsArray = new object[1];
			paramsArray.SetValue(cancel, 0);
			EventBinding.RaiseCustomEvent("Delete", ref paramsArray);

			cancel = ToInt16(paramsArray[0]);
		}

		public void BeforeDelConfirm([In] [Out] ref object cancel, [In] [Out] ref object response)
		{
            if (!Validate("BeforeDelConfirm"))
            {
                Invoker.ReleaseParamsArray(cancel, response);
                return;
            }

			object[] paramsArray = new object[2];
			paramsArray.SetValue(cancel, 0);
			paramsArray.SetValue(response, 1);
			EventBinding.RaiseCustomEvent("BeforeDelConfirm", ref paramsArray);

			cancel = ToInt16(paramsArray[0]);
            response = ToInt16(paramsArray[1]);
        }

		public void AfterDelConfirm([In] [Out] ref object status)
        {
            if (!Validate("AfterDelConfirm"))
            {
                Invoker.ReleaseParamsArray(status);
                return;
            }

			object[] paramsArray = new object[1];
			paramsArray.SetValue(status, 0);
			EventBinding.RaiseCustomEvent("AfterDelConfirm", ref paramsArray);

			status = ToInt16(paramsArray[0]);
        }

		public void Open([In] [Out] ref object cancel)
		{
            if (!Validate("Open"))
            {
                Invoker.ReleaseParamsArray(cancel);
                return;
            }

			object[] paramsArray = new object[1];
			paramsArray.SetValue(cancel, 0);
			EventBinding.RaiseCustomEvent("Open", ref paramsArray);

			cancel = ToInt16(paramsArray[0]);
        }

		public void Resize()
        {
            if (!Validate("Resize"))
            {
                return;
            }

			object[] paramsArray = new object[0];
			EventBinding.RaiseCustomEvent("Resize", ref paramsArray);
		}

		public void Unload([In] [Out] ref object cancel)
        {
            if (!Validate("Unload"))
            {
                Invoker.ReleaseParamsArray(cancel);
                return;
            }

			object[] paramsArray = new object[1];
			paramsArray.SetValue(cancel, 0);
			EventBinding.RaiseCustomEvent("Unload", ref paramsArray);

			cancel = ToInt16(paramsArray[0]);
        }

		public void Close()
        {
            if (!Validate("Close"))
            {
                return;
            }

			object[] paramsArray = new object[0];
			EventBinding.RaiseCustomEvent("Close", ref paramsArray);
		}

		public void Activate()
        {
            if (!Validate("Activate"))
            {
                return;
            }

			object[] paramsArray = new object[0];
			EventBinding.RaiseCustomEvent("Activate", ref paramsArray);
		}

		public void Deactivate()
		{
            if (!Validate("Deactivate"))
            {
                return;
            }

			object[] paramsArray = new object[0];
			EventBinding.RaiseCustomEvent("Deactivate", ref paramsArray);
		}

		public void GotFocus()
        {
            if (!Validate("GotFocus"))
            {
                return;
            }

			object[] paramsArray = new object[0];
			EventBinding.RaiseCustomEvent("GotFocus", ref paramsArray);
		}

		public void LostFocus()
        {
            if (!Validate("LostFocus"))
            {
                return;
            }

			object[] paramsArray = new object[0];
			EventBinding.RaiseCustomEvent("LostFocus", ref paramsArray);
		}

		public void Click()
		{
            if (!Validate("Click"))
            {
                return;
            }

			object[] paramsArray = new object[0];
			EventBinding.RaiseCustomEvent("Click", ref paramsArray);
		}

		public void DblClick([In] [Out] ref object cancel)
		{
            if (!Validate("DblClick"))
            {
                Invoker.ReleaseParamsArray(cancel);
                return;
            }

			object[] paramsArray = new object[1];
			paramsArray.SetValue(cancel, 0);
			EventBinding.RaiseCustomEvent("DblClick", ref paramsArray);

			cancel = ToInt16(paramsArray[0]);
		}

		public void MouseDown([In] [Out] ref object button, [In] [Out] ref object shift, [In] [Out] ref object x, [In] [Out] ref object y)
        {
            if (!Validate("MouseDown"))
            {
                Invoker.ReleaseParamsArray(button, shift, x, y);
                return;
            }

			object[] paramsArray = new object[4];
			paramsArray.SetValue(button, 0);
			paramsArray.SetValue(shift, 1);
			paramsArray.SetValue(x, 2);
			paramsArray.SetValue(y, 3);
			EventBinding.RaiseCustomEvent("MouseDown", ref paramsArray);

			button = ToInt16(paramsArray[0]);
			shift = ToInt16(paramsArray[1]);
			x = ToSingle(paramsArray[2]);
            y = ToSingle(paramsArray[3]);
        }

		public void MouseMove([In] [Out] ref object button, [In] [Out] ref object shift, [In] [Out] ref object x, [In] [Out] ref object y)
        {
            if (!Validate("MouseMove"))
            {
                Invoker.ReleaseParamsArray(button, shift, x, y);
                return;
            }

			object[] paramsArray = new object[4];
			paramsArray.SetValue(button, 0);
			paramsArray.SetValue(shift, 1);
			paramsArray.SetValue(x, 2);
			paramsArray.SetValue(y, 3);
			EventBinding.RaiseCustomEvent("MouseMove", ref paramsArray);

            button = ToInt16(paramsArray[0]);
            shift = ToInt16(paramsArray[1]);
            x = ToSingle(paramsArray[2]);
            y = ToSingle(paramsArray[3]);
        }

		public void MouseUp([In] [Out] ref object button, [In] [Out] ref object shift, [In] [Out] ref object x, [In] [Out] ref object y)
		{
            if (!Validate("MouseUp"))
            {
                Invoker.ReleaseParamsArray(button, shift, x, y);
                return;
            }

			object[] paramsArray = new object[4];
			paramsArray.SetValue(button, 0);
			paramsArray.SetValue(shift, 1);
			paramsArray.SetValue(x, 2);
			paramsArray.SetValue(y, 3);
			EventBinding.RaiseCustomEvent("MouseUp", ref paramsArray);

            button = ToInt16(paramsArray[0]);
            shift = ToInt16(paramsArray[1]);
            x = ToSingle(paramsArray[2]);
            y = ToSingle(paramsArray[3]);
        }

		public void KeyDown([In] [Out] ref object keyCode, [In] [Out] ref object shift)
        {
            if (!Validate("KeyDown"))
            {
                Invoker.ReleaseParamsArray(keyCode, shift);
                return;
            }

			object[] paramsArray = new object[2];
			paramsArray.SetValue(keyCode, 0);
			paramsArray.SetValue(shift, 1);
			EventBinding.RaiseCustomEvent("KeyDown", ref paramsArray);

			keyCode = ToInt16(paramsArray[0]);
            shift = ToInt16(paramsArray[1]);
        }

		public void KeyPress([In] [Out] ref object keyAscii)
        {
            if (!Validate("KeyPress"))
            {
                Invoker.ReleaseParamsArray(keyAscii);
                return;
            }

			object[] paramsArray = new object[1];
			paramsArray.SetValue(keyAscii, 0);
			EventBinding.RaiseCustomEvent("KeyPress", ref paramsArray);

			keyAscii = ToInt16(paramsArray[0]);
        }

		public void KeyUp([In] [Out] ref object keyCode, [In] [Out] ref object shift)
        {
            if (!Validate("KeyUp"))
            {
                Invoker.ReleaseParamsArray(keyCode, shift);
                return;
            }

			object[] paramsArray = new object[2];
			paramsArray.SetValue(keyCode, 0);
			paramsArray.SetValue(shift, 1);
			EventBinding.RaiseCustomEvent("KeyUp", ref paramsArray);

			keyCode = ToInt16(paramsArray[0]);
            shift = ToInt16(paramsArray[1]);
        }

		public void Error([In] [Out] ref object dataErr, [In] [Out] ref object response)
        {
            if (!Validate("Error"))
            {
                Invoker.ReleaseParamsArray(dataErr, response);
                return;
            }

			object[] paramsArray = new object[2];
			paramsArray.SetValue(dataErr, 0);
			paramsArray.SetValue(response, 1);
			EventBinding.RaiseCustomEvent("Error", ref paramsArray);

			dataErr = ToInt16(paramsArray[0]);
            response = ToInt16(paramsArray[1]);
        }

		public void Timer()
		{
            if (!Validate("Timer"))
            {
                return;
            }
			object[] paramsArray = new object[0];
			EventBinding.RaiseCustomEvent("Timer", ref paramsArray);
		}

		public void Filter([In] [Out] ref object cancel, [In] [Out] ref object filterType)
        {
            if (!Validate("Filter"))
            {
                Invoker.ReleaseParamsArray(cancel, filterType);
                return;
            }

			object[] paramsArray = new object[2];
			paramsArray.SetValue(cancel, 0);
			paramsArray.SetValue(filterType, 1);
			EventBinding.RaiseCustomEvent("Filter", ref paramsArray);

			cancel = ToInt16(paramsArray[0]);
            filterType = ToInt16(paramsArray[1]);
        }

		public void ApplyFilter([In] [Out] ref object cancel, [In] [Out] ref object applyType)
		{
            if (!Validate("ApplyFilter"))
            {
                Invoker.ReleaseParamsArray(cancel, applyType);
                return;
            }

			object[] paramsArray = new object[2];
			paramsArray.SetValue(cancel, 0);
			paramsArray.SetValue(applyType, 1);
			EventBinding.RaiseCustomEvent("ApplyFilter", ref paramsArray);

			cancel = ToInt16(paramsArray[0]);
            applyType = ToInt16(paramsArray[1]);
        }

		public void Dirty([In] [Out] ref object cancel)
        {
            if (!Validate("Dirty"))
            {
                Invoker.ReleaseParamsArray(cancel);
                return;
            }

			object[] paramsArray = new object[1];
			paramsArray.SetValue(cancel, 0);
			EventBinding.RaiseCustomEvent("Dirty", ref paramsArray);

			cancel = ToInt16(paramsArray[0]);
        }

		public void Undo([In] [Out] ref object cancel)
        {
            if (!Validate("Undo"))
            {
                Invoker.ReleaseParamsArray(cancel);
                return;
            }

			object[] paramsArray = new object[1];
			paramsArray.SetValue(cancel, 0);
			EventBinding.RaiseCustomEvent("Undo", ref paramsArray);

            cancel = ToInt16(paramsArray[0]);
        }

		public void RecordExit([In] [Out] ref object cancel)
        {
            if (!Validate("RecordExit"))
            {
                Invoker.ReleaseParamsArray(cancel);
                return;
            }

			object[] paramsArray = new object[1];
			paramsArray.SetValue(cancel, 0);
			EventBinding.RaiseCustomEvent("RecordExit", ref paramsArray);

            cancel = ToInt16(paramsArray[0]);
        }

		public void BeginBatchEdit([In] [Out] ref object cancel)
        {
            if (!Validate("BeginBatchEdit"))
            {
                Invoker.ReleaseParamsArray(cancel);
                return;
            }

			object[] paramsArray = new object[1];
			paramsArray.SetValue(cancel, 0);
			EventBinding.RaiseCustomEvent("BeginBatchEdit", ref paramsArray);

            cancel = ToInt16(paramsArray[0]);
        }

		public void UndoBatchEdit([In] [Out] ref object cancel)
        {
            if (!Validate("UndoBatchEdit"))
            {
                Invoker.ReleaseParamsArray(cancel);
                return;
            }

			object[] paramsArray = new object[1];
			paramsArray.SetValue(cancel, 0);
			EventBinding.RaiseCustomEvent("UndoBatchEdit", ref paramsArray);

            cancel = ToInt16(paramsArray[0]);
        }

		public void BeforeBeginTransaction([In] [Out] ref object cancel, [In] [Out, MarshalAs(UnmanagedType.IDispatch)] object connection)
		{
            if (!Validate("BeforeBeginTransaction"))
            {
                Invoker.ReleaseParamsArray(cancel, connection);
                return;
            }

            NetOffice.ADODBApi.Connection newConnection = Factory.CreateKnownObjectFromComProxy<NetOffice.ADODBApi.Connection>(EventClass, connection, NetOffice.ADODBApi.Connection.LateBindingApiWrapperType);
            (newConnection as ICOMProxyShareProvider).GetProxyShare().Acquire();

            object[] paramsArray = new object[2];
			paramsArray.SetValue(cancel, 0);
			paramsArray.SetValue(newConnection, 1);            
            EventBinding.RaiseCustomEvent("BeforeBeginTransaction", ref paramsArray);

			cancel = ToInt16(paramsArray[0]);
            connection = newConnection.UnderlyingObject;
            (newConnection as ICOMProxyShareProvider).GetProxyShare().Release();
        }

		public void AfterBeginTransaction([In] [Out, MarshalAs(UnmanagedType.IDispatch)] object connection)
		{
            if (!Validate("AfterBeginTransaction"))
            {
                Invoker.ReleaseParamsArray(connection);
                return;
            }

            NetOffice.ADODBApi.Connection newConnection = Factory.CreateKnownObjectFromComProxy<NetOffice.ADODBApi.Connection>(EventClass, connection, NetOffice.ADODBApi.Connection.LateBindingApiWrapperType);
            (newConnection as ICOMProxyShareProvider).GetProxyShare().Acquire();

            object[] paramsArray = new object[1];
			paramsArray.SetValue(newConnection, 0);
			EventBinding.RaiseCustomEvent("AfterBeginTransaction", ref paramsArray);

            connection = newConnection.UnderlyingObject;
            (newConnection as ICOMProxyShareProvider).GetProxyShare().Release();
        }

		public void BeforeCommitTransaction([In] [Out] ref object cancel, [In] [Out, MarshalAs(UnmanagedType.IDispatch)] object connection)
		{
            if (!Validate("BeforeCommitTransaction"))
            {
                Invoker.ReleaseParamsArray(connection);
                return;
            }

            NetOffice.ADODBApi.Connection newConnection = Factory.CreateKnownObjectFromComProxy<NetOffice.ADODBApi.Connection>(EventClass, connection, NetOffice.ADODBApi.Connection.LateBindingApiWrapperType);
            (newConnection as ICOMProxyShareProvider).GetProxyShare().Acquire();

            object[] paramsArray = new object[2];
			paramsArray.SetValue(cancel, 0);
			paramsArray.SetValue(newConnection, 1);
			EventBinding.RaiseCustomEvent("BeforeCommitTransaction", ref paramsArray);

			cancel = ToInt16(paramsArray[0]);
            connection = newConnection.UnderlyingObject;
            (newConnection as ICOMProxyShareProvider).GetProxyShare().Release();
        }

		public void AfterCommitTransaction([In] [Out, MarshalAs(UnmanagedType.IDispatch)] object connection)
        {
            if (!Validate("AfterCommitTransaction"))
            {
                Invoker.ReleaseParamsArray(connection);
                return;
            }

            NetOffice.ADODBApi.Connection newConnection = Factory.CreateKnownObjectFromComProxy<NetOffice.ADODBApi.Connection>(EventClass, connection, NetOffice.ADODBApi.Connection.LateBindingApiWrapperType);
            (newConnection as ICOMProxyShareProvider).GetProxyShare().Acquire();

            object[] paramsArray = new object[1];
			paramsArray.SetValue(connection, 0);
			EventBinding.RaiseCustomEvent("AfterCommitTransaction", ref paramsArray);

            connection = newConnection.UnderlyingObject;
            (newConnection as ICOMProxyShareProvider).GetProxyShare().Release();
        }

		public void RollbackTransaction([In] [Out, MarshalAs(UnmanagedType.IDispatch)] object connection)
        {
            if (!Validate("RollbackTransaction"))
            {
                Invoker.ReleaseParamsArray(connection);
                return;
            }

            NetOffice.ADODBApi.Connection newConnection = Factory.CreateKnownObjectFromComProxy<NetOffice.ADODBApi.Connection>(EventClass, connection, NetOffice.ADODBApi.Connection.LateBindingApiWrapperType);
            (newConnection as ICOMProxyShareProvider).GetProxyShare().Acquire();

            object[] paramsArray = new object[1];
			paramsArray.SetValue(connection, 0);
			EventBinding.RaiseCustomEvent("RollbackTransaction", ref paramsArray);

            connection = newConnection.UnderlyingObject;
            (newConnection as ICOMProxyShareProvider).GetProxyShare().Release();
        }

		public void OnConnect()
		{
            if (!Validate("OnConnect"))
            {
                return;
            }

			object[] paramsArray = new object[0];
			EventBinding.RaiseCustomEvent("OnConnect", ref paramsArray);
		}

		public void OnDisconnect()
		{
            if (!Validate("OnDisconnect"))
            {
                return;
            }

			object[] paramsArray = new object[0];
			EventBinding.RaiseCustomEvent("OnDisconnect", ref paramsArray);
		}

		public void PivotTableChange([In] object reason)
        {
            if (!Validate("PivotTableChange"))
            {
                return;
            }
         
			Int32 newReason = ToInt32(reason);
			object[] paramsArray = new object[1];
			paramsArray[0] = newReason;
			EventBinding.RaiseCustomEvent("PivotTableChange", ref paramsArray);
		}

		public void Query()
        {
            if (!Validate("Query"))
            {
                return;
            }

			object[] paramsArray = new object[0];
			EventBinding.RaiseCustomEvent("Query", ref paramsArray);
		}

		public void BeforeQuery()
		{
            if (!Validate("BeforeQuery"))
            {
                return;
            }

            object[] paramsArray = new object[0];
			EventBinding.RaiseCustomEvent("BeforeQuery", ref paramsArray);
		}

		public void SelectionChange()
		{
            if (!Validate("SelectionChange"))
            {
                return;
            }

            object[] paramsArray = new object[0];
			EventBinding.RaiseCustomEvent("SelectionChange", ref paramsArray);
		}

		public void CommandBeforeExecute([In] object command, [In, MarshalAs(UnmanagedType.IDispatch)] object cancel)
        {
            if (!Validate("CommandBeforeExecute"))
            {
                Invoker.ReleaseParamsArray(command, cancel);
                return;
            }

			object newCommand = (object)command;
			object newCancel = Factory.CreateEventArgumentObjectFromComProxy(EventClass, cancel) as object;
			object[] paramsArray = new object[2];
			paramsArray[0] = newCommand;
			paramsArray[1] = newCancel;
			EventBinding.RaiseCustomEvent("CommandBeforeExecute", ref paramsArray);
		}

		public void CommandChecked([In] object command, [In, MarshalAs(UnmanagedType.IDispatch)] object _checked)
        {
            if (!Validate("CommandChecked"))
            {
                Invoker.ReleaseParamsArray(command, _checked);
                return;
            }

			object newCommand = (object)command;
			object newChecked = Factory.CreateEventArgumentObjectFromComProxy(EventClass, _checked) as object;
			object[] paramsArray = new object[2];
			paramsArray[0] = newCommand;
			paramsArray[1] = newChecked;
			EventBinding.RaiseCustomEvent("CommandChecked", ref paramsArray);
		}

		public void CommandEnabled([In] object command, [In, MarshalAs(UnmanagedType.IDispatch)] object enabled)
        {
            if (!Validate("CommandEnabled"))
            {
                Invoker.ReleaseParamsArray(command, enabled);
                return;
            }

			object newCommand = (object)command;
			object newEnabled = Factory.CreateEventArgumentObjectFromComProxy(EventClass, enabled) as object;
			object[] paramsArray = new object[2];
			paramsArray[0] = newCommand;
			paramsArray[1] = newEnabled;
			EventBinding.RaiseCustomEvent("CommandEnabled", ref paramsArray);
		}

		public void CommandExecute([In] object command)
		{
            if (!Validate("CommandExecute"))
            {
                Invoker.ReleaseParamsArray(command);
                return;
            }

			object newCommand = (object)command;
			object[] paramsArray = new object[1];
			paramsArray[0] = newCommand;
            EventBinding.RaiseCustomEvent("CommandExecute", ref paramsArray);
		}

		public void DataSetChange()
        {
            if (!Validate("DataSetChange"))
            {
                return;
            }

			object[] paramsArray = new object[0];
			EventBinding.RaiseCustomEvent("DataSetChange", ref paramsArray);
		}

		public void BeforeScreenTip([In, MarshalAs(UnmanagedType.IDispatch)] object screenTipText, [In, MarshalAs(UnmanagedType.IDispatch)] object sourceObject)
        {
            if (!Validate("BeforeScreenTip"))
            {
                Invoker.ReleaseParamsArray(screenTipText, sourceObject);
                return;
            }

			object newScreenTipText = Factory.CreateEventArgumentObjectFromComProxy(EventClass, screenTipText) as object;
			object newSourceObject = Factory.CreateEventArgumentObjectFromComProxy(EventClass, sourceObject) as object;
			object[] paramsArray = new object[2];
			paramsArray[0] = newScreenTipText;
			paramsArray[1] = newSourceObject;
			EventBinding.RaiseCustomEvent("BeforeScreenTip", ref paramsArray);
		}

		public void BeforeRender([In, MarshalAs(UnmanagedType.IDispatch)] object drawObject, [In, MarshalAs(UnmanagedType.IDispatch)] object chartObject, [In, MarshalAs(UnmanagedType.IDispatch)] object cancel)
        {
            if (!Validate("BeforeRender"))
            {
                Invoker.ReleaseParamsArray(drawObject, chartObject, cancel);
                return;
            }

			object newdrawObject = Factory.CreateEventArgumentObjectFromComProxy(EventClass, drawObject) as object;
			object newchartObject = Factory.CreateEventArgumentObjectFromComProxy(EventClass, chartObject) as object;
			object newCancel = Factory.CreateEventArgumentObjectFromComProxy(EventClass, cancel) as object;
			object[] paramsArray = new object[3];
			paramsArray[0] = newdrawObject;
			paramsArray[1] = newchartObject;
			paramsArray[2] = newCancel;
			EventBinding.RaiseCustomEvent("BeforeRender", ref paramsArray);
		}

		public void AfterRender([In, MarshalAs(UnmanagedType.IDispatch)] object drawObject, [In, MarshalAs(UnmanagedType.IDispatch)] object chartObject)
        {
            if (!Validate("AfterRender"))
            {
                Invoker.ReleaseParamsArray(drawObject, chartObject);
                return;
            }

			object newdrawObject = Factory.CreateEventArgumentObjectFromComProxy(EventClass, drawObject) as object;
			object newchartObject = Factory.CreateEventArgumentObjectFromComProxy(EventClass, chartObject) as object;
			object[] paramsArray = new object[2];
			paramsArray[0] = newdrawObject;
			paramsArray[1] = newchartObject;
            EventBinding.RaiseCustomEvent("AfterRender", ref paramsArray);
		}

		public void AfterFinalRender([In, MarshalAs(UnmanagedType.IDispatch)] object drawObject)
        {
            if (!Validate("AfterFinalRender"))
            {
                Invoker.ReleaseParamsArray(drawObject);
                return;
            }

			object newdrawObject = Factory.CreateEventArgumentObjectFromComProxy(EventClass, drawObject) as object;
			object[] paramsArray = new object[1];
			paramsArray[0] = newdrawObject;
            EventBinding.RaiseCustomEvent("AfterFinalRender", ref paramsArray);
		}

		public void AfterLayout([In, MarshalAs(UnmanagedType.IDispatch)] object drawObject)
        {
            if (!Validate("AfterLayout"))
            {
                Invoker.ReleaseParamsArray(drawObject);
                return;
            }

			object newdrawObject = Factory.CreateEventArgumentObjectFromComProxy(EventClass, drawObject) as object;
			object[] paramsArray = new object[1];
			paramsArray[0] = newdrawObject;
			EventBinding.RaiseCustomEvent("AfterLayout", ref paramsArray);
		}

		public void MouseWheel([In] object page, [In] object count)
		{
            if (!Validate("MouseWheel"))
            {
                Invoker.ReleaseParamsArray(page, count);
                return;
            }

			bool newPage = ToBoolean(page);
			Int32 newCount = ToInt32(count);
			object[] paramsArray = new object[2];
			paramsArray[0] = newPage;
			paramsArray[1] = newCount;
			EventBinding.RaiseCustomEvent("MouseWheel", ref paramsArray);
		}

		public void ViewChange([In] object reason)
        {
            if (!Validate("ViewChange"))
            {
                Invoker.ReleaseParamsArray(reason);
                return;
            }
       
			Int32 newReason = ToInt32(reason);
			object[] paramsArray = new object[1];
			paramsArray[0] = newReason;
			EventBinding.RaiseCustomEvent("ViewChange", ref paramsArray);
		}

		public void DataChange([In] object reason)
		{
            if (!Validate("DataChange"))
            {
                Invoker.ReleaseParamsArray(reason);
                return;
            }

			Int32 newReason = ToInt32(reason);
			object[] paramsArray = new object[1];
			paramsArray[0] = newReason;
			EventBinding.RaiseCustomEvent("DataChange", ref paramsArray);
		}

		#endregion
	}
	
	#endregion
	
	#pragma warning restore
}