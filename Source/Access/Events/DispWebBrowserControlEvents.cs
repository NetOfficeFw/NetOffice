using System;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
using NetOffice;
using NetOffice.Attributes;

namespace NetOffice.AccessApi.EventContracts
{
    /// <summary>
    /// DispWebBrowserControlEvents
    /// </summary>
	[SupportByVersion("Access", 14,15,16)]
    [InternalEntity(InternalEntityKind.ComEventInterface)]
    [ComImport, Guid("EACB9075-68F8-4E3B-B865-E1CE6BE0447C"), InterfaceType(ComInterfaceType.InterfaceIsIDispatch), TypeLibType((short)0x1010)]
	public interface DispWebBrowserControlEvents
	{
        /// <summary>
        /// Updated
        /// </summary>
        /// <param name="code"></param>
		[SupportByVersion("Access", 14,15,16)]
        [SinkArgument("code", SinkArgumentType.Int16)]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(2076)]
		void Updated([In] [Out] ref object code);

        /// <summary>
        /// BeforeUpdate
        /// </summary>
        /// <param name="cancel"></param>
		[SupportByVersion("Access", 14,15,16)]
        [SinkArgument("code", SinkArgumentType.Int16)]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(2061)]
		void BeforeUpdate([In] [Out] ref object cancel);

        /// <summary>
        /// AfterUpdate
        /// </summary>
		[SupportByVersion("Access", 14,15,16)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(2062)]
		void AfterUpdate();

        /// <summary>
        /// Enter
        /// </summary>
		[SupportByVersion("Access", 14,15,16)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(2019)]
		void Enter();

        /// <summary>
        /// Exit
        /// </summary>
        /// <param name="cancel"></param>
		[SupportByVersion("Access", 14,15,16)]
        [SinkArgument("code", SinkArgumentType.Int16)]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(2075)]
		void Exit([In] [Out] ref object cancel);

        /// <summary>
        /// Dirty
        /// </summary>
        /// <param name="cancel"></param>
		[SupportByVersion("Access", 14,15,16)]
        [SinkArgument("code", SinkArgumentType.Int16)]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(2205)]
		void Dirty([In] [Out] ref object cancel);

        /// <summary>
        /// Change
        /// </summary>
		[SupportByVersion("Access", 14,15,16)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(2077)]
		void Change();

        /// <summary>
        /// GotFocus
        /// </summary>
		[SupportByVersion("Access", 14,15,16)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(2073)]
		void GotFocus();

        /// <summary>
        /// LostFocus
        /// </summary>
		[SupportByVersion("Access", 14,15,16)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(2074)]
		void LostFocus();

        /// <summary>
        /// Click
        /// </summary>
		[SupportByVersion("Access", 14,15,16)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(-600)]
		void Click();

        /// <summary>
        /// DblClick
        /// </summary>
        /// <param name="cancel"></param>
		[SupportByVersion("Access", 14,15,16)]
        [SinkArgument("code", SinkArgumentType.Int16)]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(-601)]
		void DblClick([In] [Out] ref object cancel);

        /// <summary>
        /// MouseDown
        /// </summary>
        /// <param name="button"></param>
        /// <param name="shift"></param>
        /// <param name="x"></param>
        /// <param name="y"></param>
		[SupportByVersion("Access", 14,15,16)]
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
		[SupportByVersion("Access", 14,15,16)]
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
        [SupportByVersion("Access", 14,15,16)]
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
		[SupportByVersion("Access", 14,15,16)]
        [SinkArgument("keyCode", SinkArgumentType.Int16)]
        [SinkArgument("shift", SinkArgumentType.Int16)]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(-602)]
		void KeyDown([In] [Out] ref object keyCode, [In] [Out] ref object shift);

        /// <summary>
        /// KeyPress
        /// </summary>
        /// <param name="keyAscii"></param>
		[SupportByVersion("Access", 14,15,16)]
        [SinkArgument("keyAscii", SinkArgumentType.Int16)]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(-603)]
		void KeyPress([In] [Out] ref object keyAscii);

        /// <summary>
        /// KeyUp
        /// </summary>
        /// <param name="keyCode"></param>
        /// <param name="shift"></param>
		[SupportByVersion("Access", 14,15,16)]
        [SinkArgument("keyCode", SinkArgumentType.Int16)]
        [SinkArgument("shift", SinkArgumentType.Int16)]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(-604)]
		void KeyUp([In] [Out] ref object keyCode, [In] [Out] ref object shift);

        /// <summary>
        /// BeforeNavigate2
        /// </summary>
        /// <param name="pDisp"></param>
        /// <param name="uRL"></param>
        /// <param name="flags"></param>
        /// <param name="targetFrameName"></param>
        /// <param name="postData"></param>
        /// <param name="headers"></param>
        /// <param name="cancel"></param>
		[SupportByVersion("Access", 14,15,16)]
        [SinkArgument("pDisp", SinkArgumentType.UnknownProxy)]
        [SinkArgument("cancel", SinkArgumentType.Bool)]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(2524)]
		void BeforeNavigate2([In, MarshalAs(UnmanagedType.IDispatch)] object pDisp, [In] [Out] ref object uRL, [In] [Out] ref object flags, [In] [Out] ref object targetFrameName, [In] [Out] ref object postData, [In] [Out] ref object headers, [In] [Out] ref object cancel);

        /// <summary>
        /// DocumentComplete
        /// </summary>
        /// <param name="pDisp"></param>
        /// <param name="uRL"></param>
		[SupportByVersion("Access", 14,15,16)]
        [SinkArgument("pDisp", SinkArgumentType.UnknownProxy)]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(2528)]
		void DocumentComplete([In, MarshalAs(UnmanagedType.IDispatch)] object pDisp, [In] [Out] ref object uRL);

        /// <summary>
        /// ProgressChange
        /// </summary>
        /// <param name="progress"></param>
        /// <param name="progressMax"></param>
		[SupportByVersion("Access", 14,15,16)]
        [SinkArgument("progress", SinkArgumentType.Int32)]
        [SinkArgument("progressMax", SinkArgumentType.Int32)]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(2515)]
		void ProgressChange([In] object progress, [In] object progressMax);

        /// <summary>
        /// NavigateError
        /// </summary>
        /// <param name="pDisp"></param>
        /// <param name="uRL"></param>
        /// <param name="targetFrameName"></param>
        /// <param name="statusCode"></param>
        /// <param name="cancel"></param>
		[SupportByVersion("Access", 14,15,16)]
        [SinkArgument("pDisp", SinkArgumentType.UnknownProxy)]
        [SinkArgument("cancel", SinkArgumentType.Bool)]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(2510)]
		void NavigateError([In, MarshalAs(UnmanagedType.IDispatch)] object pDisp, [In] [Out] ref object uRL, [In] [Out] ref object targetFrameName, [In] [Out] ref object statusCode, [In] [Out] ref object cancel);
	}
}
