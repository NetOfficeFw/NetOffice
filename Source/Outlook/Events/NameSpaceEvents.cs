using System;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
using NetOffice;
using NetOffice.Attributes;

namespace NetOffice.OutlookApi.EventContracts
{
    /// <summary>
    /// NameSpaceEvents
    /// </summary>
	[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
    [InternalEntity(InternalEntityKind.ComEventInterface)]
    [ComImport, Guid("0006308C-0000-0000-C000-000000000046"), InterfaceType(ComInterfaceType.InterfaceIsIDispatch), TypeLibType((short)0x1010)]
	public interface NameSpaceEvents
	{
        /// <summary>
        /// OptionsPagesAdd
        /// </summary>
        /// <param name="pages"></param>
        /// <param name="folder"></param>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
        [SinkArgument("pages", typeof(OutlookApi.PropertyPages))]
        [SinkArgument("newFolder", typeof(OutlookApi.MAPIFolder))]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(61445)]
		void OptionsPagesAdd([In, MarshalAs(UnmanagedType.IDispatch)] object pages, [In, MarshalAs(UnmanagedType.IDispatch)] object folder);

        /// <summary>
        /// AutoDiscoverComplete
        /// </summary>
		[SupportByVersion("Outlook", 12,14,15,16)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(64557)]
		void AutoDiscoverComplete();
	}
}
