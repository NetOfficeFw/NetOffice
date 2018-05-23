using System;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
using NetOffice;
using NetOffice.Attributes;

namespace NetOffice.VBIDEApi.EventInterfaces
{
    /// <summary>
    /// 
    /// </summary>
	[SupportByVersion("VBIDE", 12,14,5.3)]
    [InternalEntity(InternalEntityKind.ComEventInterface)]
    [ComImport, Guid("CDDE3804-2064-11CF-867F-00AA005FF34A"), InterfaceType(ComInterfaceType.InterfaceIsIDispatch), TypeLibType((short)0x1010)]
	public interface _dispReferences_Events
	{
        /// <summary>
        /// 
        /// </summary>
        /// <param name="reference"></param>
		[SupportByVersion("VBIDE", 12,14,5.3)]
        [SinkArgument("reference", typeof(NetOffice.VBIDEApi.Reference))]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(0)]
		void ItemAdded([In, MarshalAs(UnmanagedType.IDispatch)] object reference);

        /// <summary>
        /// 
        /// </summary>
        /// <param name="reference"></param>
		[SupportByVersion("VBIDE", 12,14,5.3)]
        [SinkArgument("reference", typeof(NetOffice.VBIDEApi.Reference))]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(1)]
		void ItemRemoved([In, MarshalAs(UnmanagedType.IDispatch)] object reference);
	}
}
