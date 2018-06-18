using System;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
using NetOffice;
using NetOffice.Attributes;

namespace NetOffice.VisioApi.EventContracts
{
    /// <summary>
    /// EWindows
    /// </summary>
	[SupportByVersion("Visio", 11,12,14,15,16)]
    [InternalEntity(InternalEntityKind.ComEventInterface)]
    [ComImport, Guid("000D0B01-0000-0000-C000-000000000046"), InterfaceType(ComInterfaceType.InterfaceIsIDispatch), TypeLibType((short)0x1010)]
	public interface EWindows
	{
		/// <summary>
		/// WindowOpened
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
        [SinkArgument("window", typeof(VisioApi.IVWindow))]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(32769)]
		void WindowOpened([In, MarshalAs(UnmanagedType.IDispatch)] object window);

		/// <summary>
		/// SelectionChanged
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
        [SinkArgument("window", typeof(VisioApi.IVWindow))]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(701)]
		void SelectionChanged([In, MarshalAs(UnmanagedType.IDispatch)] object window);

		/// <summary>
		/// BeforeWindowClosed
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
        [SinkArgument("window", typeof(VisioApi.IVWindow))]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(16385)]
		void BeforeWindowClosed([In, MarshalAs(UnmanagedType.IDispatch)] object window);

		/// <summary>
		/// WindowActivated
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
        [SinkArgument("window", typeof(VisioApi.IVWindow))]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(4224)]
		void WindowActivated([In, MarshalAs(UnmanagedType.IDispatch)] object window);

		/// <summary>
		/// BeforeWindowSelDelete
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
        [SinkArgument("window", typeof(VisioApi.IVWindow))]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(702)]
		void BeforeWindowSelDelete([In, MarshalAs(UnmanagedType.IDispatch)] object window);

		/// <summary>
		/// BeforeWindowPageTurn
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
        [SinkArgument("window", typeof(VisioApi.IVWindow))]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(703)]
		void BeforeWindowPageTurn([In, MarshalAs(UnmanagedType.IDispatch)] object window);

		/// <summary>
		/// WindowTurnedToPage
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
        [SinkArgument("window", typeof(VisioApi.IVWindow))]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(704)]
		void WindowTurnedToPage([In, MarshalAs(UnmanagedType.IDispatch)] object window);

		/// <summary>
		/// WindowChanged
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
        [SinkArgument("window", typeof(VisioApi.IVWindow))]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(8193)]
		void WindowChanged([In, MarshalAs(UnmanagedType.IDispatch)] object window);

		/// <summary>
		/// ViewChanged
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
        [SinkArgument("window", typeof(VisioApi.IVWindow))]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(705)]
		void ViewChanged([In, MarshalAs(UnmanagedType.IDispatch)] object window);

		/// <summary>
		/// QueryCancelWindowClose
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
        [SinkArgument("window", typeof(VisioApi.IVWindow))]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(706)]
		void QueryCancelWindowClose([In, MarshalAs(UnmanagedType.IDispatch)] object window);

		/// <summary>
		/// WindowCloseCanceled
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
        [SinkArgument("window", typeof(VisioApi.IVWindow))]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(707)]
		void WindowCloseCanceled([In, MarshalAs(UnmanagedType.IDispatch)] object window);

		/// <summary>
		/// OnKeystrokeMessageForAddon
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
        [SinkArgument("mSG", typeof(VisioApi.IVMSGWrap))]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(708)]
		void OnKeystrokeMessageForAddon([In, MarshalAs(UnmanagedType.IDispatch)] object mSG);

		/// <summary>
		/// MouseDown
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
        [SinkArgument("button", SinkArgumentType.Int32)]
        [SinkArgument("keyButtonState", SinkArgumentType.Int32)]
        [SinkArgument("x", SinkArgumentType.Double)]
        [SinkArgument("y", SinkArgumentType.Double)]
        [SinkArgument("cancelDefault", SinkArgumentType.Bool)]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(709)]
		void MouseDown([In] object button, [In] object keyButtonState, [In] object x, [In] object y, [In] [Out] ref object cancelDefault);

		/// <summary>
		/// MouseMove
		/// </summary>
        [SupportByVersion("Visio", 11,12,14,15,16)]
        [SinkArgument("button", SinkArgumentType.Int32)]
        [SinkArgument("keyButtonState", SinkArgumentType.Int32)]
        [SinkArgument("x", SinkArgumentType.Double)]
        [SinkArgument("y", SinkArgumentType.Double)]
        [SinkArgument("cancelDefault", SinkArgumentType.Bool)]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(710)]
		void MouseMove([In] object button, [In] object keyButtonState, [In] object x, [In] object y, [In] [Out] ref object cancelDefault);

		/// <summary>
		/// MouseUp
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
        [SinkArgument("button", SinkArgumentType.Int32)]
        [SinkArgument("keyButtonState", SinkArgumentType.Int32)]
        [SinkArgument("x", SinkArgumentType.Double)]
        [SinkArgument("y", SinkArgumentType.Double)]
        [SinkArgument("cancelDefault", SinkArgumentType.Bool)]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(711)]
		void MouseUp([In] object button, [In] object keyButtonState, [In] object x, [In] object y, [In] [Out] ref object cancelDefault);

		/// <summary>
		/// KeyDown
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
        [SinkArgument("keyCode", SinkArgumentType.Int32)]
        [SinkArgument("keyButtonState", SinkArgumentType.Int32)]
        [SinkArgument("cancelDefault", SinkArgumentType.Bool)]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(712)]
		void KeyDown([In] object keyCode, [In] object keyButtonState, [In] [Out] ref object cancelDefault);

		/// <summary>
		/// KeyPress
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
        [SinkArgument("keyAscii", SinkArgumentType.Int32)]
        [SinkArgument("cancelDefault", SinkArgumentType.Bool)]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(713)]
		void KeyPress([In] object keyAscii, [In] [Out] ref object cancelDefault);

		/// <summary>
		/// KeyUp
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
        [SinkArgument("keyCode", SinkArgumentType.Int32)]
        [SinkArgument("keyButtonState", SinkArgumentType.Int32)]
        [SinkArgument("cancelDefault", SinkArgumentType.Bool)]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(714)]
		void KeyUp([In] object keyCode, [In] object keyButtonState, [In] [Out] ref object cancelDefault);
	}

}
