using System;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
using NetOffice;
using NetOffice.Attributes;

namespace NetOffice.MSComctlLibApi.EventContracts
{
    /// <summary>
    /// IStatusBarEvents
    /// </summary>
    [SupportByVersion("MSComctlLib", 6)]
    [InternalEntity(InternalEntityKind.ComEventInterface)]
    [ComImport, Guid("8E3867A2-8586-11D1-B16A-00C0F0283628"), InterfaceType(ComInterfaceType.InterfaceIsIDispatch), TypeLibType((short)0x1010)]
	public interface IStatusBarEvents
	{
        /// <summary>
        /// PanelClick
        /// </summary>
        /// <param name="panel"></param>
		[SupportByVersion("MSComctlLib", 6)]
        [SinkArgument("panel", typeof(MSComctlLibApi.Panel))]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(1)]
		void PanelClick([In, MarshalAs(UnmanagedType.IDispatch)] object panel);

        /// <summary>
        /// PanelDblClick
        /// </summary>
        /// <param name="panel"></param>
		[SupportByVersion("MSComctlLib", 6)]
        [SinkArgument("panel", typeof(MSComctlLibApi.Panel))]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(2)]
		void PanelDblClick([In, MarshalAs(UnmanagedType.IDispatch)] object panel);

        /// <summary>
        /// MouseDown
        /// </summary>
        /// <param name="button"></param>
        /// <param name="shift"></param>
        /// <param name="x"></param>
        /// <param name="y"></param>
		[SupportByVersion("MSComctlLib", 6)]
        [SinkArgument("button", SinkArgumentType.Int32)]
        [SinkArgument("shift", SinkArgumentType.Int32)]
        [SinkArgument("x", SinkArgumentType.Int32)]
        [SinkArgument("y", SinkArgumentType.Int32)]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(-605)]
		void MouseDown([In] object button, [In] object shift, [In] object x, [In] object y);

        /// <summary>
        /// MouseMove
        /// </summary>
        /// <param name="button"></param>
        /// <param name="shift"></param>
        /// <param name="x"></param>
        /// <param name="y"></param>
		[SupportByVersion("MSComctlLib", 6)]
        [SinkArgument("button", SinkArgumentType.Int32)]
        [SinkArgument("shift", SinkArgumentType.Int32)]
        [SinkArgument("x", SinkArgumentType.Int32)]
        [SinkArgument("y", SinkArgumentType.Int32)]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(-606)]
		void MouseMove([In] object button, [In] object shift, [In] object x, [In] object y);

        /// <summary>
        /// MouseUp
        /// </summary>
        /// <param name="button"></param>
        /// <param name="shift"></param>
        /// <param name="x"></param>
        /// <param name="y"></param>
		[SupportByVersion("MSComctlLib", 6)]
        [SinkArgument("button", SinkArgumentType.Int32)]
        [SinkArgument("shift", SinkArgumentType.Int32)]
        [SinkArgument("x", SinkArgumentType.Int32)]
        [SinkArgument("y", SinkArgumentType.Int32)]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(-607)]
		void MouseUp([In] object button, [In] object shift, [In] object x, [In] object y);

        /// <summary>
        /// Click
        /// </summary>
		[SupportByVersion("MSComctlLib", 6)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(-600)]
		void Click();

        /// <summary>
        /// DblClick
        /// </summary>
		[SupportByVersion("MSComctlLib", 6)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(-601)]
		void DblClick();

        /// <summary>
        /// OLEStartDrag
        /// </summary>
        /// <param name="data"></param>
        /// <param name="allowedEffects"></param>
		[SupportByVersion("MSComctlLib", 6)]
        [SinkArgument("data", typeof(MSComctlLibApi.DataObject))]
        [SinkArgument("allowedEffects", SinkArgumentType.Int32)]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(1550)]
		void OLEStartDrag([In] [Out, MarshalAs(UnmanagedType.IDispatch)] object data, [In] [Out] ref object allowedEffects);

        /// <summary>
        /// OLEGiveFeedback
        /// </summary>
        /// <param name="effect"></param>
        /// <param name="defaultCursors"></param>
		[SupportByVersion("MSComctlLib", 6)]
        [SinkArgument("effect", SinkArgumentType.Int32)]
        [SinkArgument("defaultCursors", SinkArgumentType.Bool)]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(1551)]
		void OLEGiveFeedback([In] [Out] ref object effect, [In] [Out] ref object defaultCursors);

        /// <summary>
        /// OLESetData
        /// </summary>
        /// <param name="data"></param>
        /// <param name="dataFormat"></param>
		[SupportByVersion("MSComctlLib", 6)]
        [SinkArgument("data", typeof(MSComctlLibApi.DataObject))]
        [SinkArgument("dataFormat", SinkArgumentType.Int16)]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(1552)]
		void OLESetData([In] [Out, MarshalAs(UnmanagedType.IDispatch)] object data, [In] [Out] ref object dataFormat);

        /// <summary>
        /// OLECompleteDrag
        /// </summary>
        /// <param name="effect"></param>
		[SupportByVersion("MSComctlLib", 6)]
        [SinkArgument("effect", SinkArgumentType.Int32)]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(1553)]
		void OLECompleteDrag([In] [Out] ref object effect);

        /// <summary>
        /// OLEDragOver
        /// </summary>
        /// <param name="data"></param>
        /// <param name="effect"></param>
        /// <param name="button"></param>
        /// <param name="shift"></param>
        /// <param name="x"></param>
        /// <param name="y"></param>
        /// <param name="state"></param>
		[SupportByVersion("MSComctlLib", 6)]
        [SinkArgument("data", typeof(MSComctlLibApi.DataObject))]
        [SinkArgument("effect", SinkArgumentType.Int32)]
        [SinkArgument("button", SinkArgumentType.Int16)]
        [SinkArgument("shift", SinkArgumentType.Int16)]
        [SinkArgument("x", SinkArgumentType.Single)]
        [SinkArgument("y", SinkArgumentType.Single)]
        [SinkArgument("state", SinkArgumentType.Int16)]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(1554)]
		void OLEDragOver([In] [Out, MarshalAs(UnmanagedType.IDispatch)] object data, [In] [Out] ref object effect, [In] [Out] ref object button, [In] [Out] ref object shift, [In] [Out] ref object x, [In] [Out] ref object y, [In] [Out] ref object state);

        /// <summary>
        /// OLEDragDrop
        /// </summary>
        /// <param name="data"></param>
        /// <param name="effect"></param>
        /// <param name="button"></param>
        /// <param name="shift"></param>
        /// <param name="x"></param>
        /// <param name="y"></param>
		[SupportByVersion("MSComctlLib", 6)]
        [SinkArgument("data", typeof(MSComctlLibApi.DataObject))]
        [SinkArgument("effect", SinkArgumentType.Int32)]
        [SinkArgument("button", SinkArgumentType.Int16)]
        [SinkArgument("shift", SinkArgumentType.Int16)]
        [SinkArgument("x", SinkArgumentType.Single)]
        [SinkArgument("y", SinkArgumentType.Single)]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(1555)]
		void OLEDragDrop([In] [Out, MarshalAs(UnmanagedType.IDispatch)] object data, [In] [Out] ref object effect, [In] [Out] ref object button, [In] [Out] ref object shift, [In] [Out] ref object x, [In] [Out] ref object y);
	}
}
