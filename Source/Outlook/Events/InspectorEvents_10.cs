using System;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
using NetOffice;
using NetOffice.Attributes;

namespace NetOffice.OutlookApi.EventContracts
{
    /// <summary>
    /// InspectorEvents_10
    /// </summary>
	[SupportByVersion("Outlook", 10,11,12,14,15,16)]
    [InternalEntity(InternalEntityKind.ComEventInterface)]
    [ComImport, Guid("0006302A-0000-0000-C000-000000000046"), InterfaceType(ComInterfaceType.InterfaceIsIDispatch), TypeLibType((short)0x1010)]
	public interface InspectorEvents_10
	{
        /// <summary>
        /// Activate
        /// </summary>
		[SupportByVersion("Outlook", 10,11,12,14,15,16)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(61441)]
		void Activate();

        /// <summary>
        /// Deactivate
        /// </summary>
		[SupportByVersion("Outlook", 10,11,12,14,15,16)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(61446)]
		void Deactivate();

        /// <summary>
        /// Close
        /// </summary>
		[SupportByVersion("Outlook", 10,11,12,14,15,16)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(61448)]
		void Close();

        /// <summary>
        /// BeforeMaximize
        /// </summary>
        /// <param name="cancel"></param>
		[SupportByVersion("Outlook", 10,11,12,14,15,16)]
        [SinkArgument("cancel", SinkArgumentType.Bool)]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(64017)]
		void BeforeMaximize([In] [Out] ref object cancel);

        /// <summary>
        /// BeforeMinimize
        /// </summary>
        /// <param name="cancel"></param>
		[SupportByVersion("Outlook", 10,11,12,14,15,16)]
        [SinkArgument("cancel", SinkArgumentType.Bool)]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(64018)]
		void BeforeMinimize([In] [Out] ref object cancel);

        /// <summary>
        /// BeforeMove
        /// </summary>
        /// <param name="cancel"></param>
		[SupportByVersion("Outlook", 10,11,12,14,15,16)]
        [SinkArgument("cancel", SinkArgumentType.Bool)]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(64019)]
		void BeforeMove([In] [Out] ref object cancel);

        /// <summary>
        /// BeforeSize
        /// </summary>
        /// <param name="cancel"></param>
		[SupportByVersion("Outlook", 10,11,12,14,15,16)]
        [SinkArgument("cancel", SinkArgumentType.Bool)]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(64020)]
		void BeforeSize([In] [Out] ref object cancel);

        /// <summary>
        /// PageChange
        /// </summary>
        /// <param name="activePageName"></param>
		[SupportByVersion("Outlook", 12,14,15,16)]
        [SinkArgument("activePageName", SinkArgumentType.String)]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(64500)]
		void PageChange([In] [Out] ref object activePageName);

        /// <summary>
        /// AttachmentSelectionChange
        /// </summary>
		[SupportByVersion("Outlook", 14,15,16)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(64633)]
		void AttachmentSelectionChange();
	}
}
