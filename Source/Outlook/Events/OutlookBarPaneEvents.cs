using System;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
using NetOffice;
using NetOffice.Attributes;

namespace NetOffice.OutlookApi.EventContracts
{
    /// <summary>
    /// OutlookBarPaneEvents
    /// </summary>
	[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
    [InternalEntity(InternalEntityKind.ComEventInterface)]
    [ComImport, Guid("0006307A-0000-0000-C000-000000000046"), InterfaceType(ComInterfaceType.InterfaceIsIDispatch), TypeLibType((short)0x1010)]
	public interface OutlookBarPaneEvents
	{
        /// <summary>
        /// BeforeNavigate
        /// </summary>
        /// <param name="shortcut"></param>
        /// <param name="cancel"></param>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
        [SinkArgument("shortcut", typeof(OutlookApi.OutlookBarShortcut))]
        [SinkArgument("cancel", SinkArgumentType.Bool)]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(61441)]
		void BeforeNavigate([In, MarshalAs(UnmanagedType.IDispatch)] object shortcut, [In] [Out] ref object cancel);

        /// <summary>
        /// BeforeGroupSwitch
        /// </summary>
        /// <param name="toGroup"></param>
        /// <param name="cancel"></param>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
        [SinkArgument("toGroup", typeof(OutlookApi.OutlookBarGroup))]
        [SinkArgument("cancel", SinkArgumentType.Bool)]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(61442)]
		void BeforeGroupSwitch([In, MarshalAs(UnmanagedType.IDispatch)] object toGroup, [In] [Out] ref object cancel);
	}
}
