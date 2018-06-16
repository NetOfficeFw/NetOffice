using System;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
using NetOffice;
using NetOffice.Attributes;

namespace NetOffice.OWC10Api.EventContracts
{
    /// <summary>
    /// IChartEvents
    /// </summary>
    [SupportByVersion("OWC10", 1)]
    [InternalEntity(InternalEntityKind.ComEventInterface)]
    [ComImport, Guid("F5B39A7A-1480-11D3-8549-00C04FAC67D7"), InterfaceType(ComInterfaceType.InterfaceIsIDispatch), TypeLibType((short)0x1010)]
	public interface IChartEvents
	{
        /// <summary>
        /// DataSetChange
        /// </summary>
		[SupportByVersion("OWC10", 1)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(5101)]
		void DataSetChange();

        /// <summary>
        /// DblClick
        /// </summary>
		[SupportByVersion("OWC10", 1)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(5102)]
		void DblClick();

        /// <summary>
        /// Click
        /// </summary>
		[SupportByVersion("OWC10", 1)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(5103)]
		void Click();

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
        /// KeyPress
        /// </summary>
        /// <param name="keyAscii"></param>
		[SupportByVersion("OWC10", 1)]
        [SinkArgument("keyAscii", SinkArgumentType.Int32)]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(1010)]
		void KeyPress([In] object keyAscii);

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
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(5107)]
		void MouseDown([In] object button, [In] object shift, [In] object x, [In] object y);

        /// <summary>
        /// MouseMove
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
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(5108)]
		void MouseMove([In] object button, [In] object shift, [In] object x, [In] object y);

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
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(5109)]
		void MouseUp([In] object button, [In] object shift, [In] object x, [In] object y);

        /// <summary>
        /// MouseWheel
        /// </summary>
        /// <param name="page"></param>
        /// <param name="count"></param>
		[SupportByVersion("OWC10", 1)]
        [SinkArgument("page", SinkArgumentType.Bool)]
        [SinkArgument("count", SinkArgumentType.Int32)]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(5118)]
		void MouseWheel([In] object page, [In] object count);

        /// <summary>
        /// SelectionChange
        /// </summary>
		[SupportByVersion("OWC10", 1)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(5110)]
		void SelectionChange();

        /// <summary>
        /// BeforeScreenTip
        /// </summary>
        /// <param name="tipText"></param>
        /// <param name="contextObject"></param>
		[SupportByVersion("OWC10", 1)]
        [SinkArgument("tipText", typeof(OWC10Api.ByRef))]
        [SinkArgument("newContextObject", SinkArgumentType.UnknownProxy)]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(5120)]
		void BeforeScreenTip([In, MarshalAs(UnmanagedType.IDispatch)] object tipText, [In, MarshalAs(UnmanagedType.IDispatch)] object contextObject);

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
        [SinkArgument("succeeded", SinkArgumentType.Bool)]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(1004)]
		void CommandExecute([In] object command, [In] object succeeded);

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
        /// BeforeRender
        /// </summary>
        /// <param name="drawObject"></param>
        /// <param name="chartObject"></param>
        /// <param name="cancel"></param>
		[SupportByVersion("OWC10", 1)]
        [SinkArgument("drawObject", typeof(OWC10Api.ChChartDraw))]
        [SinkArgument("chartObject", SinkArgumentType.UnknownProxy)]
        [SinkArgument("cancel", typeof(OWC10Api.ByRef))]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(5111)]
		void BeforeRender([In, MarshalAs(UnmanagedType.IDispatch)] object drawObject, [In, MarshalAs(UnmanagedType.IDispatch)] object chartObject, [In, MarshalAs(UnmanagedType.IDispatch)] object cancel);

        /// <summary>
        /// AfterRender
        /// </summary>
        /// <param name="drawObject"></param>
        /// <param name="chartObject"></param>
		[SupportByVersion("OWC10", 1)]
        [SinkArgument("drawObject", typeof(OWC10Api.ChChartDraw))]
        [SinkArgument("chartObject", SinkArgumentType.UnknownProxy)]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(5112)]
		void AfterRender([In, MarshalAs(UnmanagedType.IDispatch)] object drawObject, [In, MarshalAs(UnmanagedType.IDispatch)] object chartObject);

        /// <summary>
        /// AfterFinalRender
        /// </summary>
        /// <param name="drawObject"></param>
		[SupportByVersion("OWC10", 1)]
        [SinkArgument("drawObject", typeof(OWC10Api.ChChartDraw))]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(5113)]
		void AfterFinalRender([In, MarshalAs(UnmanagedType.IDispatch)] object drawObject);

        /// <summary>
        /// AfterLayout
        /// </summary>
        /// <param name="drawObject"></param>
		[SupportByVersion("OWC10", 1)]
        [SinkArgument("drawObject", typeof(OWC10Api.ChChartDraw))]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(5114)]
		void AfterLayout([In, MarshalAs(UnmanagedType.IDispatch)] object drawObject);

        /// <summary>
        /// ViewChange
        /// </summary>
		[SupportByVersion("OWC10", 1)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(5119)]
		void ViewChange();
	}
}
