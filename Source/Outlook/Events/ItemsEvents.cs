using System;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
using NetOffice;
using NetOffice.Attributes;

namespace NetOffice.OutlookApi.EventContracts
{
    /// <summary>
    /// ItemsEvents
    /// </summary>
	[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
    [InternalEntity(InternalEntityKind.ComEventInterface)]
    [ComImport, Guid("00063077-0000-0000-C000-000000000046"), InterfaceType(ComInterfaceType.InterfaceIsIDispatch), TypeLibType((short)0x1010)]
	public interface ItemsEvents
	{
        /// <summary>
        /// ItemAdd
        /// </summary>
        /// <param name="item"></param>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
        [SinkArgument("item", SinkArgumentType.UnknownProxy)]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(61441)]
		void ItemAdd([In, MarshalAs(UnmanagedType.IDispatch)] object item);

        /// <summary>
        /// ItemChange
        /// </summary>
        /// <param name="item"></param>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
        [SinkArgument("item", SinkArgumentType.UnknownProxy)]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(61442)]
		void ItemChange([In, MarshalAs(UnmanagedType.IDispatch)] object item);

        /// <summary>
        /// ItemRemove
        /// </summary>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(61443)]
		void ItemRemove();
	}
}
