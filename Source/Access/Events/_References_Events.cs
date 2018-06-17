using System;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
using NetOffice;
using NetOffice.Attributes;

namespace NetOffice.AccessApi.EventContracts
{
    /// <summary>
    /// _References_Events
    /// </summary>
	[SupportByVersion("Access", 9,10,11,12,14,15,16)]
    [InternalEntity(InternalEntityKind.ComEventInterface)]
    [ComImport, Guid("F163F201-ADA2-11CF-89A9-00A0C9054129"), InterfaceType(ComInterfaceType.InterfaceIsIDispatch), TypeLibType((short)0x1010)]
	public interface _References_Events
	{
        /// <summary>
        /// ItemAdded
        /// </summary>
        /// <param name="reference"></param>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
        [SinkArgument("reference", typeof(NetOffice.AccessApi.Reference))]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(0)]
		void ItemAdded([In, MarshalAs(UnmanagedType.IDispatch)] object reference);

        /// <summary>
        /// ItemRemoved
        /// </summary>
        /// <param name="reference"></param>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
        [SinkArgument("reference", typeof(NetOffice.AccessApi.Reference))]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(1)]
		void ItemRemoved([In, MarshalAs(UnmanagedType.IDispatch)] object reference);
	}
}
