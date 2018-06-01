using System;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
using NetOffice;
using NetOffice.Attributes;

namespace NetOffice.OfficeApi.EventContracts
{
    /// <summary>
    /// _CustomTaskPaneEvents
    /// </summary>
    [SupportByVersion("Office", 12,14,15,16)]
    [InternalEntity(InternalEntityKind.ComEventInterface)]
    [ComImport, Guid("000C033C-0000-0000-C000-000000000046"), InterfaceType(ComInterfaceType.InterfaceIsIDispatch), TypeLibType((short)0x1010)]
	public interface _CustomTaskPaneEvents
	{
        /// <summary>
        /// VisibleStateChange
        /// </summary>
        /// <param name="customTaskPaneInst"></param>
		[SupportByVersion("Office", 12,14,15,16)]
        [SinkArgument("customTaskPaneInst", typeof(NetOffice.OfficeApi._CustomTaskPane))]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(1)]
		void VisibleStateChange([In, MarshalAs(UnmanagedType.IDispatch)] object customTaskPaneInst);

        /// <summary>
        /// DockPositionStateChange
        /// </summary>
        /// <param name="customTaskPaneInst"></param>
		[SupportByVersion("Office", 12,14,15,16)]
        [SinkArgument("customTaskPaneInst", typeof(NetOffice.OfficeApi._CustomTaskPane))]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(2)]
		void DockPositionStateChange([In, MarshalAs(UnmanagedType.IDispatch)] object customTaskPaneInst);
	}
}
