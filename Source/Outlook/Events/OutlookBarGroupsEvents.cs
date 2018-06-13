using System;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
using NetOffice;
using NetOffice.Attributes;

namespace NetOffice.OutlookApi.EventContracts
{
    /// <summary>
    /// OutlookBarGroupsEvents
    /// </summary>
	[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
    [InternalEntity(InternalEntityKind.ComEventInterface)]
    [ComImport, Guid("0006307B-0000-0000-C000-000000000046"), InterfaceType(ComInterfaceType.InterfaceIsIDispatch), TypeLibType((short)0x1010)]
	public interface OutlookBarGroupsEvents
	{
        /// <summary>
        /// GroupAdd
        /// </summary>
        /// <param name="newGroup"></param>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
        [SinkArgument("newGroup", typeof(OutlookApi.OutlookBarGroup))]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(61441)]
		void GroupAdd([In, MarshalAs(UnmanagedType.IDispatch)] object newGroup);

        /// <summary>
        /// BeforeGroupAdd
        /// </summary>
        /// <param name="cancel"></param>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
        [SinkArgument("cancel", SinkArgumentType.Bool)]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(61442)]
		void BeforeGroupAdd([In] [Out] ref object cancel);

        /// <summary>
        /// BeforeGroupRemove
        /// </summary>
        /// <param name="group"></param>
        /// <param name="cancel"></param>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
        [SinkArgument("newGroup", typeof(OutlookApi.OutlookBarGroup))]
        [SinkArgument("cancel", SinkArgumentType.Bool)]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(61443)]
		void BeforeGroupRemove([In, MarshalAs(UnmanagedType.IDispatch)] object group, [In] [Out] ref object cancel);
	}
}
