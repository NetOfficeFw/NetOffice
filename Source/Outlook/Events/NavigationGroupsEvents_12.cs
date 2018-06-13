using System;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
using NetOffice;
using NetOffice.Attributes;

namespace NetOffice.OutlookApi.EventContracts
{
    /// <summary>
    /// NavigationGroupsEvents_12
    /// </summary>
	[SupportByVersion("Outlook", 12,14,15,16)]
    [InternalEntity(InternalEntityKind.ComEventInterface)]
    [ComImport, Guid("000630F4-0000-0000-C000-000000000046"), InterfaceType(ComInterfaceType.InterfaceIsIDispatch), TypeLibType((short)0x1010)]
	public interface NavigationGroupsEvents_12
	{
        /// <summary>
        /// SelectedChange
        /// </summary>
        /// <param name="navigationFolder"></param>
		[SupportByVersion("Outlook", 12,14,15,16)]
        [SinkArgument("navigationFolder", typeof(OutlookApi.NavigationFolder))]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(64458)]
		void SelectedChange([In, MarshalAs(UnmanagedType.IDispatch)] object navigationFolder);

        /// <summary>
        /// NavigationFolderAdd
        /// </summary>
        /// <param name="navigationFolder"></param>
		[SupportByVersion("Outlook", 12,14,15,16)]
        [SinkArgument("navigationFolder", typeof(OutlookApi.NavigationFolder))]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(64459)]
		void NavigationFolderAdd([In, MarshalAs(UnmanagedType.IDispatch)] object navigationFolder);

        /// <summary>
        /// NavigationFolderRemove
        /// </summary>
		[SupportByVersion("Outlook", 12,14,15,16)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(64460)]
		void NavigationFolderRemove();
	}
}
