using System;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
using NetOffice;
using NetOffice.Attributes;

namespace NetOffice.OutlookApi.EventContracts
{
    /// <summary>
    /// OutlookBarShortcutsEvents
    /// </summary>
	[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
    [InternalEntity(InternalEntityKind.ComEventInterface)]
    [ComImport, Guid("0006307C-0000-0000-C000-000000000046"), InterfaceType(ComInterfaceType.InterfaceIsIDispatch), TypeLibType((short)0x1010)]
	public interface OutlookBarShortcutsEvents
	{
        /// <summary>
        /// ShortcutAdd
        /// </summary>
        /// <param name="newShortcut"></param>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
        [SinkArgument("newShortcut", typeof(OutlookApi.OutlookBarShortcut))]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(61441)]
		void ShortcutAdd([In, MarshalAs(UnmanagedType.IDispatch)] object newShortcut);

        /// <summary>
        /// BeforeShortcutAdd
        /// </summary>
        /// <param name="cancel"></param>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
        [SinkArgument("cancel", SinkArgumentType.Bool)]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(61442)]
		void BeforeShortcutAdd([In] [Out] ref object cancel);

        /// <summary>
        /// BeforeShortcutRemove
        /// </summary>
        /// <param name="shortcut"></param>
        /// <param name="cancel"></param>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
        [SinkArgument("shortcut", typeof(OutlookApi.OutlookBarShortcut))]
        [SinkArgument("cancel", SinkArgumentType.Bool)]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(61443)]
		void BeforeShortcutRemove([In, MarshalAs(UnmanagedType.IDispatch)] object shortcut, [In] [Out] ref object cancel);
	}
}
