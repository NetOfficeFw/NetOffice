using System;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
using NetOffice;
using NetOffice.Attributes;

namespace NetOffice.OWC10Api.EventContracts
{
    /// <summary>
    /// ISpreadsheetEventSink
    /// </summary>
    [SupportByVersion("OWC10", 1)]
    [InternalEntity(InternalEntityKind.ComEventInterface)]
    [ComImport, Guid("F5B39A75-1480-11D3-8549-00C04FAC67D7"), InterfaceType(ComInterfaceType.InterfaceIsIDispatch), TypeLibType((short)0x1010)]
	public interface ISpreadsheetEventSink
	{
        /// <summary>
        /// BeforeContextMenu
        /// </summary>
        /// <param name="x"></param>
        /// <param name="y"></param>
        /// <param name="menu"></param>
        /// <param name="cancel"></param>
		[SupportByVersion("OWC10", 1)]
        [SinkArgument("x", SinkArgumentType.Int32)]
        [SinkArgument("y", SinkArgumentType.Int32)]
        [SinkArgument("menu", typeof(OWC10Api.ByRef))]
        [SinkArgument("cancel", typeof(OWC10Api.ByRef))]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(1011)]
		void BeforeContextMenu([In] object x, [In] object y, [In, MarshalAs(UnmanagedType.IDispatch)] object menu, [In, MarshalAs(UnmanagedType.IDispatch)] object cancel);

        /// <summary>
        /// BeforeKeyDown
        /// </summary>
        /// <param name="keyCode"></param>
        /// <param name="shift"></param>
        /// <param name="cancel"></param>
		[SupportByVersion("OWC10", 1)]
        [SinkArgument("keyCode", SinkArgumentType.Int32)]
        [SinkArgument("shift", SinkArgumentType.Int32)]
        [SinkArgument("cancel", typeof(OWC10Api.ByRef))]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(1006)]
		void BeforeKeyDown([In] object keyCode, [In] object shift, [In, MarshalAs(UnmanagedType.IDispatch)] object cancel);

        /// <summary>
        /// BeforeKeyPress
        /// </summary>
        /// <param name="keyAscii"></param>
        /// <param name="cancel"></param>
		[SupportByVersion("OWC10", 1)]
        [SinkArgument("keyAscii", SinkArgumentType.Int32)]
        [SinkArgument("cancel", typeof(OWC10Api.ByRef))]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(1007)]
		void BeforeKeyPress([In] object keyAscii, [In, MarshalAs(UnmanagedType.IDispatch)] object cancel);

        /// <summary>
        /// BeforeKeyUp
        /// </summary>
        /// <param name="keyCode"></param>
        /// <param name="shift"></param>
        /// <param name="cancel"></param>
		[SupportByVersion("OWC10", 1)]
        [SinkArgument("keyCode", SinkArgumentType.Int32)]
        [SinkArgument("shift", SinkArgumentType.Int32)]
        [SinkArgument("cancel", typeof(OWC10Api.ByRef))]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(1005)]
		void BeforeKeyUp([In] object keyCode, [In] object shift, [In, MarshalAs(UnmanagedType.IDispatch)] object cancel);

        /// <summary>
        /// Click
        /// </summary>
		[SupportByVersion("OWC10", 1)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(1502)]
		void Click();

        /// <summary>
        /// CommandEnabled
        /// </summary>
        /// <param name="command"></param>
        /// <param name="enabled"></param>
		[SupportByVersion("OWC10", 1)]
        [SinkArgument("enabled", typeof(OWC10Api.ByRef))]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(1000)]
		void CommandEnabled([In] object command, [In, MarshalAs(UnmanagedType.IDispatch)] object enabled);

        /// <summary>
        /// CommandChecked
        /// </summary>
        /// <param name="command"></param>
        /// <param name="_checked"></param>
		[SupportByVersion("OWC10", 1)]
        [SinkArgument("_checked", typeof(OWC10Api.ByRef))]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(1001)]
		void CommandChecked([In] object command, [In, MarshalAs(UnmanagedType.IDispatch)] object _checked);

        /// <summary>
        /// CommandTipText
        /// </summary>
        /// <param name="command"></param>
        /// <param name="caption"></param>
		[SupportByVersion("OWC10", 1)]
        [SinkArgument("caption", typeof(OWC10Api.ByRef))]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(1002)]
		void CommandTipText([In] object command, [In, MarshalAs(UnmanagedType.IDispatch)] object caption);

        /// <summary>
        /// CommandBeforeExecute
        /// </summary>
        /// <param name="command"></param>
        /// <param name="cancel"></param>
		[SupportByVersion("OWC10", 1)]
        [SinkArgument("cancel", typeof(OWC10Api.ByRef))]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(1003)]
		void CommandBeforeExecute([In] object command, [In, MarshalAs(UnmanagedType.IDispatch)] object cancel);

        /// <summary>
        /// CommandExecute
        /// </summary>
        /// <param name="command"></param>
        /// <param name="succeeded"></param>
		[SupportByVersion("OWC10", 1)]
        [SinkArgument("succeeded", typeof(OWC10Api.ByRef))]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(1004)]
		void CommandExecute([In] object command, [In] object succeeded);

        /// <summary>
        /// DblClick
        /// </summary>
		[SupportByVersion("OWC10", 1)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(1503)]
		void DblClick();

        /// <summary>
        /// EndEdit
        /// </summary>
        /// <param name="accept"></param>
        /// <param name="finalValue"></param>
        /// <param name="cancel"></param>
        /// <param name="errorDescription"></param>
		[SupportByVersion("OWC10", 1)]
        [SinkArgument("accept", SinkArgumentType.Bool)]
        [SinkArgument("finalValue", typeof(OWC10Api.ByRef))]
        [SinkArgument("cancel", typeof(OWC10Api.ByRef))]
        [SinkArgument("errorDescription", typeof(OWC10Api.ByRef))]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(1504)]
		void EndEdit([In] object accept, [In, MarshalAs(UnmanagedType.IDispatch)] object finalValue, [In, MarshalAs(UnmanagedType.IDispatch)] object cancel, [In, MarshalAs(UnmanagedType.IDispatch)] object errorDescription);

        /// <summary>
        /// Initialize
        /// </summary>
		[SupportByVersion("OWC10", 1)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(1523)]
		void Initialize();

        /// <summary>
        /// KeyDown
        /// </summary>
        /// <param name="keyCode"></param>
        /// <param name="shift"></param>
		[SupportByVersion("OWC10", 1)]
        [SinkArgument("keyCode", SinkArgumentType.Int32)]
        [SinkArgument("shift", SinkArgumentType.Int32)]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(1009)]
		void KeyDown([In] object keyCode, [In] object shift);

        /// <summary>
        /// KeyPress
        /// </summary>
        /// <param name="keyAscii"></param>
		[SupportByVersion("OWC10", 1)]  
        [SinkArgument("keyAscii", SinkArgumentType.Int32)]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(1010)]
		void KeyPress([In] object keyAscii);

        /// <summary>
        /// KeyUp
        /// </summary>
        /// <param name="keyCode"></param>
        /// <param name="shift"></param>
		[SupportByVersion("OWC10", 1)]
        [SinkArgument("keyCode", SinkArgumentType.Int32)]
        [SinkArgument("shift", SinkArgumentType.Int32)]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(1008)]
		void KeyUp([In] object keyCode, [In] object shift);

        /// <summary>
        /// LoadCompleted
        /// </summary>
		[SupportByVersion("OWC10", 1)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(1522)]
		void LoadCompleted();

        /// <summary>
        /// MouseDown
        /// </summary>
        /// <param name="button"></param>
        /// <param name="shift"></param>
        /// <param name="x"></param>
        /// <param name="y"></param>
		[SupportByVersion("OWC10", 1)]
        [SinkArgument("button", SinkArgumentType.Int32)]
        [SinkArgument("shift", SinkArgumentType.Int32)]
        [SinkArgument("x", SinkArgumentType.Int32)]
        [SinkArgument("y", SinkArgumentType.Int32)]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(1505)]
		void MouseDown([In] object button, [In] object shift, [In] object x, [In] object y);

        /// <summary>
        /// MouseOut
        /// </summary>
        /// <param name="button"></param>
        /// <param name="shift"></param>
        /// <param name="target"></param>
		[SupportByVersion("OWC10", 1)]
        [SinkArgument("button", SinkArgumentType.Int32)]
        [SinkArgument("shift", SinkArgumentType.Int32)]
        [SinkArgument("target", typeof(OWC10Api._Range))]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(1506)]
		void MouseOut([In] object button, [In] object shift, [In, MarshalAs(UnmanagedType.IDispatch)] object target);

        /// <summary>
        /// MouseOver
        /// </summary>
        /// <param name="button"></param>
        /// <param name="shift"></param>
        /// <param name="target"></param>
		[SupportByVersion("OWC10", 1)]
        [SinkArgument("button", SinkArgumentType.Int32)]
        [SinkArgument("shift", SinkArgumentType.Int32)]
        [SinkArgument("target", typeof(OWC10Api._Range))] 
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(1507)]
		void MouseOver([In] object button, [In] object shift, [In, MarshalAs(UnmanagedType.IDispatch)] object target);

        /// <summary>
        /// MouseUp
        /// </summary>
        /// <param name="button"></param>
        /// <param name="shift"></param>
        /// <param name="x"></param>
        /// <param name="y"></param>
		[SupportByVersion("OWC10", 1)]
        [SinkArgument("button", SinkArgumentType.Int32)]
        [SinkArgument("shift", SinkArgumentType.Int32)]
        [SinkArgument("x", SinkArgumentType.Int32)]
        [SinkArgument("y", SinkArgumentType.Int32)]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(1508)]
		void MouseUp([In] object button, [In] object shift, [In] object x, [In] object y);

        /// <summary>
        /// MouseWheel
        /// </summary>
        /// <param name="page"></param>
        /// <param name="count"></param>
		[SupportByVersion("OWC10", 1)]
        [SinkArgument("page", SinkArgumentType.Int32)]
        [SinkArgument("count", SinkArgumentType.Bool)]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(1509)]
		void MouseWheel([In] object page, [In] object count);

        /// <summary>
        /// SelectionChange
        /// </summary>
        [SinkArgument("target", typeof(OWC10Api._Range))]
        [SupportByVersion("OWC10", 1)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(1511)]
		void SelectionChange();

        /// <summary>
        /// SelectionChanging
        /// </summary>
        /// <param name="range"></param>
		[SupportByVersion("OWC10", 1)]
        [SinkArgument("range", typeof(OWC10Api._Range))]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(1512)]
		void SelectionChanging([In, MarshalAs(UnmanagedType.IDispatch)] object range);

        /// <summary>
        /// SheetActivate
        /// </summary>
        /// <param name="sh"></param>
		[SupportByVersion("OWC10", 1)]
        [SinkArgument("sh", typeof(OWC10Api.Worksheet))]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(1513)]
		void SheetActivate([In, MarshalAs(UnmanagedType.IDispatch)] object sh);

        /// <summary>
        /// SheetCalculate
        /// </summary>
        /// <param name="sh"></param>
		[SupportByVersion("OWC10", 1)]
        [SinkArgument("sh", typeof(OWC10Api.Worksheet))]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(1516)]
		void SheetCalculate([In, MarshalAs(UnmanagedType.IDispatch)] object sh);

        /// <summary>
        /// SheetChange
        /// </summary>
        /// <param name="sh"></param>
        /// <param name="target"></param>
		[SupportByVersion("OWC10", 1)]
        [SinkArgument("sh", typeof(OWC10Api.Worksheet))]
        [SinkArgument("target", typeof(OWC10Api._Range))]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(1517)]
		void SheetChange([In, MarshalAs(UnmanagedType.IDispatch)] object sh, [In, MarshalAs(UnmanagedType.IDispatch)] object target);

        /// <summary>
        /// SheetDeactivate
        /// </summary>
        /// <param name="sh"></param>
		[SupportByVersion("OWC10", 1)]
        [SinkArgument("sh", typeof(OWC10Api.Worksheet))]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(1518)]
		void SheetDeactivate([In, MarshalAs(UnmanagedType.IDispatch)] object sh);

        /// <summary>
        /// SheetFollowHyperlink
        /// </summary>
        /// <param name="sh"></param>
        /// <param name="target"></param>
		[SupportByVersion("OWC10", 1)]
        [SinkArgument("sh", typeof(OWC10Api.Worksheet))]
        [SinkArgument("target", typeof(OWC10Api._Range))]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(1519)]
		void SheetFollowHyperlink([In, MarshalAs(UnmanagedType.IDispatch)] object sh, [In, MarshalAs(UnmanagedType.IDispatch)] object target);

        /// <summary>
        /// StartEdit
        /// </summary>
        /// <param name="selection"></param>
        /// <param name="initialValue"></param>
        /// <param name="cancel"></param>
        /// <param name="errorDescription"></param>
		[SupportByVersion("OWC10", 1)]
        [SinkArgument("selection", SinkArgumentType.UnknownProxy)]
        [SinkArgument("initialValue", typeof(OWC10Api.ByRef))]
        [SinkArgument("cancel", typeof(OWC10Api.ByRef))]
        [SinkArgument("errorDescription", typeof(OWC10Api.ByRef))]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(1520)]
		void StartEdit([In, MarshalAs(UnmanagedType.IDispatch)] object selection, [In, MarshalAs(UnmanagedType.IDispatch)] object initialValue, [In, MarshalAs(UnmanagedType.IDispatch)] object cancel, [In, MarshalAs(UnmanagedType.IDispatch)] object errorDescription);

        /// <summary>
        /// ViewChange
        /// </summary>
        /// <param name="target"></param>
		[SupportByVersion("OWC10", 1)]
        [SinkArgument("target", typeof(OWC10Api._Range))]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(1521)]
		void ViewChange([In, MarshalAs(UnmanagedType.IDispatch)] object target);
	}
}
