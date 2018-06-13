using System;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
using NetOffice;
using NetOffice.Attributes;

namespace NetOffice.OutlookApi.EventContracts
{
    /// <summary>
    /// OlkListBoxEvents
    /// </summary>
	[SupportByVersion("Outlook", 12,14,15,16)]
    [InternalEntity(InternalEntityKind.ComEventInterface)]
    [ComImport, Guid("000672E4-0000-0000-C000-000000000046"), InterfaceType(ComInterfaceType.InterfaceIsIDispatch), TypeLibType((short)0x1010)]
	public interface OlkListBoxEvents
	{
        /// <summary>
        /// Click
        /// </summary>
		[SupportByVersion("Outlook", 12,14,15,16)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(-600)]
		void Click();

        /// <summary>
        /// DoubleClick
        /// </summary>
		[SupportByVersion("Outlook", 12,14,15,16)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(-601)]
		void DoubleClick();

        /// <summary>
        /// MouseDown
        /// </summary>
        /// <param name="button"></param>
        /// <param name="shift"></param>
        /// <param name="x"></param>
        /// <param name="y"></param>
        [SupportByVersion("Outlook", 12,14,15,16)]
        [SinkArgument("button", SinkArgumentType.Enum, typeof(OutlookApi.Enums.OlMouseButton))]
        [SinkArgument("shift", SinkArgumentType.Enum, typeof(OutlookApi.Enums.OlShiftState))]
        [SinkArgument("x", SinkArgumentType.Single)]
        [SinkArgument("y", SinkArgumentType.Single)]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(-605)]
		void MouseDown([In] object button, [In] object shift, [In] object x, [In] object y);

        /// <summary>
        /// MouseMove
        /// </summary>
        /// <param name="button"></param>
        /// <param name="shift"></param>
        /// <param name="x"></param>
        /// <param name="y"></param>
		[SupportByVersion("Outlook", 12,14,15,16)]
        [SinkArgument("button", SinkArgumentType.Enum, typeof(OutlookApi.Enums.OlMouseButton))]
        [SinkArgument("shift", SinkArgumentType.Enum, typeof(OutlookApi.Enums.OlShiftState))]
        [SinkArgument("x", SinkArgumentType.Single)]
        [SinkArgument("y", SinkArgumentType.Single)]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(-606)]
		void MouseMove([In] object button, [In] object shift, [In] object x, [In] object y);

        /// <summary>
        /// MouseUp
        /// </summary>
        /// <param name="button"></param>
        /// <param name="shift"></param>
        /// <param name="x"></param>
        /// <param name="y"></param>
		[SupportByVersion("Outlook", 12,14,15,16)]
        [SinkArgument("button", SinkArgumentType.Enum, typeof(OutlookApi.Enums.OlMouseButton))]
        [SinkArgument("shift", SinkArgumentType.Enum, typeof(OutlookApi.Enums.OlShiftState))]
        [SinkArgument("x", SinkArgumentType.Single)]
        [SinkArgument("y", SinkArgumentType.Single)]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(-607)]
		void MouseUp([In] object button, [In] object shift, [In] object x, [In] object y);

        /// <summary>
        /// Enter
        /// </summary>
		[SupportByVersion("Outlook", 12,14,15,16)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(-2147384830)]
		void Enter();

        /// <summary>
        /// Exit
        /// </summary>
        /// <param name="cancel"></param>
		[SupportByVersion("Outlook", 12,14,15,16)]
        [SinkArgument("cancel", SinkArgumentType.Bool)]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(-2147384829)]
		void Exit([In] [Out] ref object cancel);

        /// <summary>
        /// KeyDown
        /// </summary>
        /// <param name="keyCode"></param>
        /// <param name="shift"></param>
		[SupportByVersion("Outlook", 12,14,15,16)]
        [SinkArgument("keyCode", SinkArgumentType.Int32)]
        [SinkArgument("shift", SinkArgumentType.Enum, typeof(OutlookApi.Enums.OlShiftState))]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(-602)]
		void KeyDown([In] [Out] ref object keyCode, [In] object shift);

        /// <summary>
        /// KeyPress
        /// </summary>
        /// <param name="keyAscii"></param>
		[SupportByVersion("Outlook", 12,14,15,16)]
        [SinkArgument("keyAscii", SinkArgumentType.Int32)]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(-603)]
		void KeyPress([In] [Out] ref object keyAscii);

        /// <summary>
        /// KeyUp
        /// </summary>
        /// <param name="keyCode"></param>
        /// <param name="shift"></param>
		[SupportByVersion("Outlook", 12,14,15,16)]
        [SinkArgument("keyCode", SinkArgumentType.Int32)]
        [SinkArgument("shift", SinkArgumentType.Enum, typeof(OutlookApi.Enums.OlShiftState))]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(-604)]
		void KeyUp([In] [Out] ref object keyCode, [In] object shift);

        /// <summary>
        /// Change
        /// </summary>
		[SupportByVersion("Outlook", 12,14,15,16)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(2)]
		void Change();

        /// <summary>
        /// AfterUpdate
        /// </summary>
		[SupportByVersion("Outlook", 12,14,15,16)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(-2147384832)]
		void AfterUpdate();

        /// <summary>
        /// BeforeUpdate
        /// </summary>
        /// <param name="cancel"></param>
		[SupportByVersion("Outlook", 12,14,15,16)]
        [SinkArgument("cancel", SinkArgumentType.Bool)]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(-2147384831)]
		void BeforeUpdate([In] [Out] ref object cancel);
	}
}
