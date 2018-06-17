using System;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
using NetOffice;
using NetOffice.Attributes;

namespace NetOffice.AccessApi.EventContracts
{
    /// <summary>
    /// _FormEvents
    /// </summary>
	[SupportByVersion("Access", 9,10,11,12,14,15,16)]
    [InternalEntity(InternalEntityKind.ComEventInterface)]
    [ComImport, Guid("331FDCFB-CF31-11CD-8701-00AA003F0F07"), InterfaceType(ComInterfaceType.InterfaceIsIDispatch), TypeLibType((short)0x1010)]
	public interface _FormEvents
	{
        /// <summary>
        /// Load
        /// </summary>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(2067)]
		void Load();

        /// <summary>
        /// Current
        /// </summary>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(2058)]
		void Current();

        /// <summary>
        /// BeforeInsert
        /// </summary>
        /// <param name="cancel"></param>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(2059)]
		void BeforeInsert([In] [Out] ref object cancel);

        /// <summary>
        /// AfterInsert
        /// </summary>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(2060)]
		void AfterInsert();

        /// <summary>
        /// BeforeUpdate
        /// </summary>
        /// <param name="cancel"></param>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(2061)]
		void BeforeUpdate([In] [Out] ref object cancel);

        /// <summary>
        /// AfterUpdate
        /// </summary>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(2062)]
		void AfterUpdate();

        /// <summary>
        /// Delete
        /// </summary>
        /// <param name="cancel"></param>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
        [SinkArgument("cancel", SinkArgumentType.Int16)]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(2063)]
		void Delete([In] [Out] ref object cancel);

        /// <summary>
        /// BeforeDelConfirm
        /// </summary>
        /// <param name="cancel"></param>
        /// <param name="response"></param>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
        [SinkArgument("cancel", SinkArgumentType.Int16)]
        [SinkArgument("response", SinkArgumentType.Int16)]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(2064)]
		void BeforeDelConfirm([In] [Out] ref object cancel, [In] [Out] ref object response);

        /// <summary>
        /// AfterDelConfirm
        /// </summary>
        /// <param name="status"></param>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
        [SinkArgument("status", SinkArgumentType.Int16)]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(2065)]
		void AfterDelConfirm([In] [Out] ref object status);

        /// <summary>
        /// Open
        /// </summary>
        /// <param name="cancel"></param>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
        [SinkArgument("cancel", SinkArgumentType.Int16)]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(2066)]
		void Open([In] [Out] ref object cancel);

        /// <summary>
        /// Resize
        /// </summary>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(2068)]
		void Resize();

        /// <summary>
        /// Unload
        /// </summary>
        /// <param name="cancel"></param>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
        [SinkArgument("cancel", SinkArgumentType.Int16)]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(2069)]
		void Unload([In] [Out] ref object cancel);

        /// <summary>
        /// Close
        /// </summary>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(2070)]
		void Close();

        /// <summary>
        /// Activate
        /// </summary>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(2071)]
		void Activate();

        /// <summary>
        /// Deactivate
        /// </summary>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(2072)]
		void Deactivate();

        /// <summary>
        /// GotFocus
        /// </summary>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(2073)]
		void GotFocus();

        /// <summary>
        /// LostFocus
        /// </summary>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(2074)]
		void LostFocus();

        /// <summary>
        /// Click
        /// </summary>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(-600)]
		void Click();

        /// <summary>
        /// DblClick
        /// </summary>
        /// <param name="cancel"></param>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
        [SinkArgument("cancel", SinkArgumentType.Int16)]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(-601)]
		void DblClick([In] [Out] ref object cancel);

        /// <summary>
        /// MouseDown
        /// </summary>
        /// <param name="button"></param>
        /// <param name="shift"></param>
        /// <param name="x"></param>
        /// <param name="y"></param>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
        [SinkArgument("button", SinkArgumentType.Int16)]
        [SinkArgument("shift", SinkArgumentType.Int16)]
        [SinkArgument("x", SinkArgumentType.Single)]
        [SinkArgument("y", SinkArgumentType.Single)]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(-605)]
		void MouseDown([In] [Out] ref object button, [In] [Out] ref object shift, [In] [Out] ref object x, [In] [Out] ref object y);

        /// <summary>
        /// MouseMove
        /// </summary>
        /// <param name="button"></param>
        /// <param name="shift"></param>
        /// <param name="x"></param>
        /// <param name="y"></param>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
        [SinkArgument("button", SinkArgumentType.Int16)]
        [SinkArgument("shift", SinkArgumentType.Int16)]
        [SinkArgument("x", SinkArgumentType.Single)]
        [SinkArgument("y", SinkArgumentType.Single)]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(-606)]
		void MouseMove([In] [Out] ref object button, [In] [Out] ref object shift, [In] [Out] ref object x, [In] [Out] ref object y);

        /// <summary>
        /// MouseUp
        /// </summary>
        /// <param name="button"></param>
        /// <param name="shift"></param>
        /// <param name="x"></param>
        /// <param name="y"></param>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
        [SinkArgument("button", SinkArgumentType.Int16)]
        [SinkArgument("shift", SinkArgumentType.Int16)]
        [SinkArgument("x", SinkArgumentType.Single)]
        [SinkArgument("y", SinkArgumentType.Single)]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(-607)]
		void MouseUp([In] [Out] ref object button, [In] [Out] ref object shift, [In] [Out] ref object x, [In] [Out] ref object y);

        /// <summary>
        /// KeyDown
        /// </summary>
        /// <param name="keyCode"></param>
        /// <param name="shift"></param>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
        [SinkArgument("keyCode", SinkArgumentType.Int16)]
        [SinkArgument("shift", SinkArgumentType.Int16)]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(-602)]
		void KeyDown([In] [Out] ref object keyCode, [In] [Out] ref object shift);

        /// <summary>
        /// KeyPress
        /// </summary>
        /// <param name="keyAscii"></param>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
        [SinkArgument("keyAscii", SinkArgumentType.Int16)]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(-603)]
		void KeyPress([In] [Out] ref object keyAscii);

        /// <summary>
        /// KeyUp
        /// </summary>
        /// <param name="keyCode"></param>
        /// <param name="shift"></param>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
        [SinkArgument("keyCode", SinkArgumentType.Int16)]
        [SinkArgument("shift", SinkArgumentType.Int16)]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(-604)]
		void KeyUp([In] [Out] ref object keyCode, [In] [Out] ref object shift);

        /// <summary>
        /// Error
        /// </summary>
        /// <param name="dataErr"></param>
        /// <param name="response"></param>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
        [SinkArgument("dataErr", SinkArgumentType.Int16)]
        [SinkArgument("response", SinkArgumentType.Int16)]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(2083)]
		void Error([In] [Out] ref object dataErr, [In] [Out] ref object response);

        /// <summary>
        /// Timer
        /// </summary>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(2084)]
		void Timer();

        /// <summary>
        /// Filter
        /// </summary>
        /// <param name="cancel"></param>
        /// <param name="filterType"></param>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
        [SinkArgument("cancel", SinkArgumentType.Int16)]
        [SinkArgument("filterType", SinkArgumentType.Int16)]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(2155)]
		void Filter([In] [Out] ref object cancel, [In] [Out] ref object filterType);

        /// <summary>
        /// ApplyFilter
        /// </summary>
        /// <param name="cancel"></param>
        /// <param name="applyType"></param>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
        [SinkArgument("cancel", SinkArgumentType.Int16)]
        [SinkArgument("filterType", SinkArgumentType.Int16)]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(2156)]
		void ApplyFilter([In] [Out] ref object cancel, [In] [Out] ref object applyType);

        /// <summary>
        /// Dirty
        /// </summary>
        /// <param name="cancel"></param>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
        [SinkArgument("cancel", SinkArgumentType.Int16)]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(2205)]
		void Dirty([In] [Out] ref object cancel);

        /// <summary>
        /// Undo
        /// </summary>
        /// <param name="cancel"></param>
		[SupportByVersion("Access", 10,11,12,14,15,16)]
        [SinkArgument("cancel", SinkArgumentType.Int16)]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(2145)]
		void Undo([In] [Out] ref object cancel);

        /// <summary>
        /// RecordExit
        /// </summary>
        /// <param name="cancel"></param>
		[SupportByVersion("Access", 10,11,12,14,15,16)]
        [SinkArgument("cancel", SinkArgumentType.Int16)]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(2334)]
		void RecordExit([In] [Out] ref object cancel);

        /// <summary>
        /// BeginBatchEdit
        /// </summary>
        /// <param name="cancel"></param>
		[SupportByVersion("Access", 10,11,12,14,15,16)]
        [SinkArgument("cancel", SinkArgumentType.Int16)]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(2369)]
		void BeginBatchEdit([In] [Out] ref object cancel);

        /// <summary>
        /// UndoBatchEdit
        /// </summary>
        /// <param name="cancel"></param>
		[SupportByVersion("Access", 10,11,12,14,15,16)]
        [SinkArgument("cancel", SinkArgumentType.Int16)]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(2370)]
		void UndoBatchEdit([In] [Out] ref object cancel);

        /// <summary>
        /// BeforeBeginTransaction
        /// </summary>
        /// <param name="cancel"></param>
        /// <param name="connection"></param>
        [SupportByVersion("Access", 10,11,12,14,15,16)]
        [SinkArgument("cancel", SinkArgumentType.Int16)]
        [SinkArgument("connection", typeof(NetOffice.ADODBApi.Connection))]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(2371)]
		void BeforeBeginTransaction([In] [Out] ref object cancel, [In] [Out, MarshalAs(UnmanagedType.IDispatch)] object connection);

        /// <summary>
        /// AfterBeginTransaction
        /// </summary>
        /// <param name="connection"></param>
		[SupportByVersion("Access", 10,11,12,14,15,16)]
        [SinkArgument("connection", typeof(NetOffice.ADODBApi.Connection))]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(2372)]
		void AfterBeginTransaction([In] [Out, MarshalAs(UnmanagedType.IDispatch)] object connection);

        /// <summary>
        /// BeforeCommitTransaction
        /// </summary>
        /// <param name="cancel"></param>
        /// <param name="connection"></param>
		[SupportByVersion("Access", 10,11,12,14,15,16)]
        [SinkArgument("cancel", SinkArgumentType.Int16)]
        [SinkArgument("connection", typeof(NetOffice.ADODBApi.Connection))]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(2373)]
		void BeforeCommitTransaction([In] [Out] ref object cancel, [In] [Out, MarshalAs(UnmanagedType.IDispatch)] object connection);

        /// <summary>
        /// AfterCommitTransaction
        /// </summary>
        /// <param name="connection"></param>
		[SupportByVersion("Access", 10,11,12,14,15,16)]
        [SinkArgument("connection", typeof(NetOffice.ADODBApi.Connection))]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(2374)]
		void AfterCommitTransaction([In] [Out, MarshalAs(UnmanagedType.IDispatch)] object connection);

        /// <summary>
        /// RollbackTransaction
        /// </summary>
        /// <param name="connection"></param>
		[SupportByVersion("Access", 10,11,12,14,15,16)]
        [SinkArgument("connection", typeof(NetOffice.ADODBApi.Connection))]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(2375)]
		void RollbackTransaction([In] [Out, MarshalAs(UnmanagedType.IDispatch)] object connection);

        /// <summary>
        /// OnConnect
        /// </summary>
		[SupportByVersion("Access", 10,11,12,14,15,16)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(2383)]
		void OnConnect();

        /// <summary>
        /// OnDisconnect
        /// </summary>
		[SupportByVersion("Access", 10,11,12,14,15,16)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(2384)]
		void OnDisconnect();

        /// <summary>
        /// PivotTableChange
        /// </summary>
        /// <param name="reason"></param>
		[SupportByVersion("Access", 10,11,12,14,15,16)]
        [SinkArgument("reason", SinkArgumentType.Int32)]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(2385)]
		void PivotTableChange([In] object reason);

        /// <summary>
        /// Query
        /// </summary>
		[SupportByVersion("Access", 10,11,12,14,15,16)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(2386)]
		void Query();

        /// <summary>
        /// BeforeQuery
        /// </summary>
		[SupportByVersion("Access", 10,11,12,14,15,16)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(2387)]
		void BeforeQuery();

        /// <summary>
        /// SelectionChange
        /// </summary>
		[SupportByVersion("Access", 10,11,12,14,15,16)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(2388)]
		void SelectionChange();

        /// <summary>
        /// CommandBeforeExecute
        /// </summary>
        /// <param name="command"></param>
        /// <param name="cancel"></param>
		[SupportByVersion("Access", 10,11,12,14,15,16)]
        [SinkArgument("cancel", SinkArgumentType.UnknownProxy)]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(2389)]
		void CommandBeforeExecute([In] object command, [In, MarshalAs(UnmanagedType.IDispatch)] object cancel);

        /// <summary>
        /// CommandChecked
        /// </summary>
        /// <param name="command"></param>
        /// <param name="_checked"></param>
		[SupportByVersion("Access", 10,11,12,14,15,16)]
        [SinkArgument("_checked", SinkArgumentType.UnknownProxy)]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(2390)]
		void CommandChecked([In] object command, [In, MarshalAs(UnmanagedType.IDispatch)] object _checked);

        /// <summary>
        /// CommandEnabled
        /// </summary>
        /// <param name="command"></param>
        /// <param name="enabled"></param>
		[SupportByVersion("Access", 10,11,12,14,15,16)]
        [SinkArgument("enabled", SinkArgumentType.UnknownProxy)]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(2391)]
		void CommandEnabled([In] object command, [In, MarshalAs(UnmanagedType.IDispatch)] object enabled);

        /// <summary>
        /// CommandExecute
        /// </summary>
        /// <param name="command"></param>
		[SupportByVersion("Access", 10,11,12,14,15,16)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(2392)]
		void CommandExecute([In] object command);

        /// <summary>
        /// DataSetChange
        /// </summary>
		[SupportByVersion("Access", 10,11,12,14,15,16)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(2394)]
		void DataSetChange();

        /// <summary>
        /// BeforeScreenTip
        /// </summary>
        /// <param name="screenTipText"></param>
        /// <param name="sourceObject"></param>
		[SupportByVersion("Access", 10,11,12,14,15,16)]
        [SinkArgument("screenTipText", SinkArgumentType.UnknownProxy)]
        [SinkArgument("sourceObject", SinkArgumentType.UnknownProxy)]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(2395)]
		void BeforeScreenTip([In, MarshalAs(UnmanagedType.IDispatch)] object screenTipText, [In, MarshalAs(UnmanagedType.IDispatch)] object sourceObject);

        /// <summary>
        /// BeforeRender
        /// </summary>
        /// <param name="drawObject"></param>
        /// <param name="chartObject"></param>
        /// <param name="cancel"></param>
		[SupportByVersion("Access", 10,11,12,14,15,16)]
        [SinkArgument("drawObject", SinkArgumentType.UnknownProxy)]
        [SinkArgument("chartObject", SinkArgumentType.UnknownProxy)]
        [SinkArgument("cancel", SinkArgumentType.UnknownProxy)]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(2399)]
		void BeforeRender([In, MarshalAs(UnmanagedType.IDispatch)] object drawObject, [In, MarshalAs(UnmanagedType.IDispatch)] object chartObject, [In, MarshalAs(UnmanagedType.IDispatch)] object cancel);

        /// <summary>
        /// AfterRender
        /// </summary>
        /// <param name="drawObject"></param>
        /// <param name="chartObject"></param>
		[SupportByVersion("Access", 10,11,12,14,15,16)]
        [SinkArgument("drawObject", SinkArgumentType.UnknownProxy)]
        [SinkArgument("chartObject", SinkArgumentType.UnknownProxy)]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(2397)]
		void AfterRender([In, MarshalAs(UnmanagedType.IDispatch)] object drawObject, [In, MarshalAs(UnmanagedType.IDispatch)] object chartObject);

        /// <summary>
        /// AfterFinalRender
        /// </summary>
        /// <param name="drawObject"></param>
		[SupportByVersion("Access", 10,11,12,14,15,16)]
        [SinkArgument("drawObject", SinkArgumentType.UnknownProxy)]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(2396)]
		void AfterFinalRender([In, MarshalAs(UnmanagedType.IDispatch)] object drawObject);

        /// <summary>
        /// AfterLayout
        /// </summary>
        /// <param name="drawObject"></param>
		[SupportByVersion("Access", 10,11,12,14,15,16)]
        [SinkArgument("drawObject", SinkArgumentType.UnknownProxy)]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(2398)]
		void AfterLayout([In, MarshalAs(UnmanagedType.IDispatch)] object drawObject);

        /// <summary>
        /// MouseWheel
        /// </summary>
        /// <param name="page"></param>
        /// <param name="count"></param>
		[SupportByVersion("Access", 10,11,12,14,15,16)]
        [SinkArgument("drawObject", SinkArgumentType.Bool)]
        [SinkArgument("count", SinkArgumentType.Int32)]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(2401)]
		void MouseWheel([In] object page, [In] object count);

        /// <summary>
        /// ViewChange
        /// </summary>
        /// <param name="reason"></param>
		[SupportByVersion("Access", 10,11,12,14,15,16)]
        [SinkArgument("reason", SinkArgumentType.Int32)]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(2402)]
		void ViewChange([In] object reason);

        /// <summary>
        /// DataChange
        /// </summary>
        /// <param name="reason"></param>
		[SupportByVersion("Access", 10,11,12,14,15,16)]
        [SinkArgument("reason", SinkArgumentType.Int32)]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(2403)]
		void DataChange([In] object reason);
	}
}
