using System;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
using NetOffice;
using NetOffice.Attributes;

namespace NetOffice.OutlookApi.EventContracts
{
    /// <summary>
    /// ExplorersEvents
    /// </summary>
	[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
    [InternalEntity(InternalEntityKind.ComEventInterface)]
    [ComImport, Guid("00063078-0000-0000-C000-000000000046"), InterfaceType(ComInterfaceType.InterfaceIsIDispatch), TypeLibType((short)0x1010)]
	public interface ExplorersEvents
	{
        /// <summary>
        /// NewExplorer
        /// </summary>
        /// <param name="explorer"></param>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
        [SinkArgument("explorer", typeof(NetOffice.OutlookApi._Explorer))]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(61441)]
		void NewExplorer([In, MarshalAs(UnmanagedType.IDispatch)] object explorer);
	}
}
