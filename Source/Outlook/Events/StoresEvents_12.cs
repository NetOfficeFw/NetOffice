using System;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
using NetOffice;
using NetOffice.Attributes;

namespace NetOffice.OutlookApi.EventContracts
{
    /// <summary>
    /// StoresEvents_12
    /// </summary>
	[SupportByVersion("Outlook", 12,14,15,16)]
    [InternalEntity(InternalEntityKind.ComEventInterface)]
    [ComImport, Guid("000630F8-0000-0000-C000-000000000046"), InterfaceType(ComInterfaceType.InterfaceIsIDispatch), TypeLibType((short)0x1010)]
	public interface StoresEvents_12
	{
        /// <summary>
        /// BeforeStoreRemove
        /// </summary>
        /// <param name="store"></param>
        /// <param name="cancel"></param>
		[SupportByVersion("Outlook", 12,14,15,16)]
        [SinkArgument("store", typeof(NetOffice.OutlookApi._Store))]
        [SinkArgument("cancel", SinkArgumentType.Bool)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(64433)]
		void BeforeStoreRemove([In, MarshalAs(UnmanagedType.IDispatch)] object store, [In] [Out] ref object cancel);

        /// <summary>
        /// StoreAdd
        /// </summary>
        /// <param name="store"></param>
		[SupportByVersion("Outlook", 12,14,15,16)]
        [SinkArgument("store", typeof(NetOffice.OutlookApi._Store))]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(61441)]
		void StoreAdd([In, MarshalAs(UnmanagedType.IDispatch)] object store);
	}
}
