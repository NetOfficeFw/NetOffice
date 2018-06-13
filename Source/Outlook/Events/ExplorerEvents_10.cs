using System;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
using NetOffice;
using NetOffice.Attributes;

namespace NetOffice.OutlookApi.EventContracts
{
    /// <summary>
    /// ExplorerEvents_10
    /// </summary>
	[SupportByVersion("Outlook", 10,11,12,14,15,16)]
    [InternalEntity(InternalEntityKind.ComEventInterface)]
    [ComImport, Guid("0006300F-0000-0000-C000-000000000046"), InterfaceType(ComInterfaceType.InterfaceIsIDispatch), TypeLibType((short)0x1010)]
	public interface ExplorerEvents_10
	{
        /// <summary>
        /// Activate
        /// </summary>
		[SupportByVersion("Outlook", 10,11,12,14,15,16)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(61441)]
		void Activate();

        /// <summary>
        /// FolderSwitch
        /// </summary>
		[SupportByVersion("Outlook", 10,11,12,14,15,16)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(61442)]
		void FolderSwitch();

        /// <summary>
        /// BeforeFolderSwitch
        /// </summary>
        /// <param name="newFolder"></param>
        /// <param name="cancel"></param>
		[SupportByVersion("Outlook", 10,11,12,14,15,16)]
        [SinkArgument("newFolder", SinkArgumentType.UnknownProxy)]
        [SinkArgument("cancel", SinkArgumentType.Bool)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(61443)]
		void BeforeFolderSwitch([In, MarshalAs(UnmanagedType.IDispatch)] object newFolder, [In] [Out] ref object cancel);

        /// <summary>
        /// ViewSwitch
        /// </summary>
		[SupportByVersion("Outlook", 10,11,12,14,15,16)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(61444)]
		void ViewSwitch();

        /// <summary>
        /// BeforeViewSwitch
        /// </summary>
        /// <param name="newView"></param>
        /// <param name="cancel"></param>
		[SupportByVersion("Outlook", 10,11,12,14,15,16)]
        [SinkArgument("cancel", SinkArgumentType.Bool)]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(61445)]
		void BeforeViewSwitch([In] object newView, [In] [Out] ref object cancel);

        /// <summary>
        /// Deactivate
        /// </summary>
		[SupportByVersion("Outlook", 10,11,12,14,15,16)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(61446)]
		void Deactivate();

        /// <summary>
        /// SelectionChange
        /// </summary>
		[SupportByVersion("Outlook", 10,11,12,14,15,16)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(61447)]
		void SelectionChange();

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
        /// BeforeItemCopy
        /// </summary>
        /// <param name="cancel"></param>
		[SupportByVersion("Outlook", 10,11,12,14,15,16)]
        [SinkArgument("cancel", SinkArgumentType.Bool)]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(64014)]
		void BeforeItemCopy([In] [Out] ref object cancel);

        /// <summary>
        /// BeforeItemCut
        /// </summary>
        /// <param name="cancel"></param>
		[SupportByVersion("Outlook", 10,11,12,14,15,16)]
        [SinkArgument("cancel", SinkArgumentType.Bool)]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(64015)]
		void BeforeItemCut([In] [Out] ref object cancel);

        /// <summary>
        /// BeforeItemPaste
        /// </summary>
        /// <param name="clipboardContent"></param>
        /// <param name="target"></param>
        /// <param name="cancel"></param>
		[SupportByVersion("Outlook", 10,11,12,14,15,16)]
        [SinkArgument("target", typeof(OutlookApi.MAPIFolder))]
        [SinkArgument("cancel", SinkArgumentType.Bool)]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(64016)]
		void BeforeItemPaste([In] [Out] ref object clipboardContent, [In, MarshalAs(UnmanagedType.IDispatch)] object target, [In] [Out] ref object cancel);

        /// <summary>
        /// AttachmentSelectionChange
        /// </summary>
		[SupportByVersion("Outlook", 14,15,16)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(64633)]
		void AttachmentSelectionChange();

        /// <summary>
        /// InlineResponse
        /// </summary>
        /// <param name="item"></param>
		[SupportByVersion("Outlook", 15, 16)]
        [SinkArgument("item", SinkArgumentType.UnknownProxy)]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(64658)]
		void InlineResponse([In, MarshalAs(UnmanagedType.IDispatch)] object item);

        /// <summary>
        /// InlineResponseClose
        /// </summary>
        [SupportByVersion("Outlook", 15, 16)]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(64662)]
        void InlineResponseClose();
	}
}
