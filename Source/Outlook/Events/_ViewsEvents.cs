using System;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
using NetOffice;
using NetOffice.Attributes;

namespace NetOffice.OutlookApi.EventContracts
{
    /// <summary>
    /// _ViewsEvents
    /// </summary>
	[SupportByVersion("Outlook", 10,11,12,14,15,16)]
    [InternalEntity(InternalEntityKind.ComEventInterface)]
    [ComImport, Guid("000630A5-0000-0000-C000-000000000046"), InterfaceType(ComInterfaceType.InterfaceIsIDispatch), TypeLibType((short)0x1010)]
	public interface _ViewsEvents
	{
        /// <summary>
        /// ViewAdd
        /// </summary>
        /// <param name="view"></param>
		[SupportByVersion("Outlook", 10,11,12,14,15,16)]
        [SinkArgument("view", typeof(OutlookApi.View))]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(400)]
		void ViewAdd([In, MarshalAs(UnmanagedType.IDispatch)] object view);

        /// <summary>
        /// ViewRemove
        /// </summary>
        /// <param name="view"></param>
		[SupportByVersion("Outlook", 10,11,12,14,15,16)]
        [SinkArgument("view", typeof(OutlookApi.View))]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(64071)]
		void ViewRemove([In, MarshalAs(UnmanagedType.IDispatch)] object view);
	}
}
