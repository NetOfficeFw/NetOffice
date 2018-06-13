using System;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
using NetOffice;
using NetOffice.Attributes;

namespace NetOffice.OutlookApi.EventContracts
{
    /// <summary>
    /// ApplicationEvents_11
    /// </summary>
	[SupportByVersion("Outlook", 11,12,14,15,16)]
    [InternalEntity(InternalEntityKind.ComEventInterface)]
    [ComImport, Guid("0006302C-0000-0000-C000-000000000046"), InterfaceType(ComInterfaceType.InterfaceIsIDispatch), TypeLibType((short)0x1010)]
	public interface ApplicationEvents_11
	{
        /// <summary>
        /// ItemSend
        /// </summary>
        /// <param name="item"></param>
        /// <param name="cancel"></param>
		[SupportByVersion("Outlook", 11,12,14,15,16)]
        [SinkArgument("item", SinkArgumentType.UnknownProxy)]
        [SinkArgument("cancel", SinkArgumentType.Bool)]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(61442)]
		void ItemSend([In, MarshalAs(UnmanagedType.IDispatch)] object item, [In] [Out] ref object cancel);

        /// <summary>
        /// NewMail
        /// </summary>
		[SupportByVersion("Outlook", 11,12,14,15,16)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(61443)]
		void NewMail();

        /// <summary>
        /// Reminder
        /// </summary>
        /// <param name="item"></param>
		[SupportByVersion("Outlook", 11,12,14,15,16)]
        [SinkArgument("item", SinkArgumentType.UnknownProxy)]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(61444)]
		void Reminder([In, MarshalAs(UnmanagedType.IDispatch)] object item);

        /// <summary>
        /// OptionsPagesAdd
        /// </summary>
        /// <param name="pages"></param>
		[SupportByVersion("Outlook", 11,12,14,15,16)]
        [SinkArgument("pages", typeof(OutlookApi.PropertyPages))]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(61445)]
		void OptionsPagesAdd([In, MarshalAs(UnmanagedType.IDispatch)] object pages);

        /// <summary>
        /// Startup
        /// </summary>
		[SupportByVersion("Outlook", 11,12,14,15,16)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(61446)]
		void Startup();

        /// <summary>
        /// Quit
        /// </summary>
		[SupportByVersion("Outlook", 11,12,14,15,16)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(61447)]
		void Quit();

        /// <summary>
        /// AdvancedSearchComplete
        /// </summary>
        /// <param name="searchObject"></param>
		[SupportByVersion("Outlook", 11,12,14,15,16)]
        [SinkArgument("searchObject", typeof(OutlookApi.Search))]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(64106)]
		void AdvancedSearchComplete([In, MarshalAs(UnmanagedType.IDispatch)] object searchObject);

        /// <summary>
        /// AdvancedSearchStopped
        /// </summary>
        /// <param name="searchObject"></param>
		[SupportByVersion("Outlook", 11,12,14,15,16)]
        [SinkArgument("searchObject", typeof(OutlookApi.Search))]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(64107)]
		void AdvancedSearchStopped([In, MarshalAs(UnmanagedType.IDispatch)] object searchObject);

        /// <summary>
        /// MAPILogonComplete
        /// </summary>
		[SupportByVersion("Outlook", 11,12,14,15,16)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(64144)]
		void MAPILogonComplete();

        /// <summary>
        /// NewMailEx
        /// </summary>
        /// <param name="entryIDCollection"></param>
		[SupportByVersion("Outlook", 11,12,14,15,16)]
        [SinkArgument("entryIDCollection", SinkArgumentType.String)]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(64181)]
		void NewMailEx([In] object entryIDCollection);

        /// <summary>
        /// AttachmentContextMenuDisplay
        /// </summary>
        /// <param name="commandBar"></param>
        /// <param name="attachments"></param>
		[SupportByVersion("Outlook", 12,14,15,16)]
        [SinkArgument("commandBar", typeof(OfficeApi.CommandBar))]
        [SinkArgument("attachments", typeof(OutlookApi.AttachmentSelection))]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(64318)]
		void AttachmentContextMenuDisplay([In, MarshalAs(UnmanagedType.IDispatch)] object commandBar, [In, MarshalAs(UnmanagedType.IDispatch)] object attachments);

        /// <summary>
        /// FolderContextMenuDisplay
        /// </summary>
        /// <param name="commandBar"></param>
        /// <param name="folder"></param>
		[SupportByVersion("Outlook", 12,14,15,16)]
        [SinkArgument("commandBar", typeof(OfficeApi.CommandBar))]
        [SinkArgument("folder", typeof(OutlookApi.Folder))]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(64322)]
		void FolderContextMenuDisplay([In, MarshalAs(UnmanagedType.IDispatch)] object commandBar, [In, MarshalAs(UnmanagedType.IDispatch)] object folder);

        /// <summary>
        /// StoreContextMenuDisplay
        /// </summary>
        /// <param name="commandBar"></param>
        /// <param name="store"></param>
		[SupportByVersion("Outlook", 12,14,15,16)]
        [SinkArgument("commandBar", typeof(OfficeApi.CommandBar))]
        [SinkArgument("store", typeof(OutlookApi.Store))]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(64323)]
		void StoreContextMenuDisplay([In, MarshalAs(UnmanagedType.IDispatch)] object commandBar, [In, MarshalAs(UnmanagedType.IDispatch)] object store);

        /// <summary>
        /// ShortcutContextMenuDisplay
        /// </summary>
        /// <param name="commandBar"></param>
        /// <param name="shortcut"></param>
		[SupportByVersion("Outlook", 12,14,15,16)]
        [SinkArgument("commandBar", typeof(OfficeApi.CommandBar))]
        [SinkArgument("shortcut", typeof(OutlookApi.OutlookBarShortcut))]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(64324)]
		void ShortcutContextMenuDisplay([In, MarshalAs(UnmanagedType.IDispatch)] object commandBar, [In, MarshalAs(UnmanagedType.IDispatch)] object shortcut);

        /// <summary>
        /// ViewContextMenuDisplay
        /// </summary>
        /// <param name="commandBar"></param>
        /// <param name="view"></param>
		[SupportByVersion("Outlook", 12,14,15,16)]
        [SinkArgument("commandBar", typeof(OfficeApi.CommandBar))]
        [SinkArgument("view", typeof(OutlookApi.View))]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(64320)]
		void ViewContextMenuDisplay([In, MarshalAs(UnmanagedType.IDispatch)] object commandBar, [In, MarshalAs(UnmanagedType.IDispatch)] object view);

        /// <summary>
        /// ItemContextMenuDisplay
        /// </summary>
        /// <param name="commandBar"></param>
        /// <param name="selection"></param>
		[SupportByVersion("Outlook", 12,14,15,16)]
        [SinkArgument("commandBar", typeof(OfficeApi.CommandBar))]
        [SinkArgument("selection", typeof(OutlookApi.Selection))]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(64321)]
		void ItemContextMenuDisplay([In, MarshalAs(UnmanagedType.IDispatch)] object commandBar, [In, MarshalAs(UnmanagedType.IDispatch)] object selection);

        /// <summary>
        /// ContextMenuClose
        /// </summary>
        /// <param name="contextMenu"></param>
		[SupportByVersion("Outlook", 12,14,15,16)]
        [SinkArgument("contextMenu", SinkArgumentType.Enum, typeof(OutlookApi.Enums.OlContextMenu))]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(64422)]
		void ContextMenuClose([In] object contextMenu);

        /// <summary>
        /// ItemLoad
        /// </summary>
        /// <param name="item"></param>
		[SupportByVersion("Outlook", 12,14,15,16)]
        [SinkArgument("item", SinkArgumentType.UnknownProxy)]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(64423)]
		void ItemLoad([In, MarshalAs(UnmanagedType.IDispatch)] object item);

        /// <summary>
        /// BeforeFolderSharingDialog
        /// </summary>
        /// <param name="folderToShare"></param>
        /// <param name="cancel"></param>
		[SupportByVersion("Outlook", 12,14,15,16)]
        [SinkArgument("folderToShare", typeof(OutlookApi.MAPIFolder))]
        [SinkArgument("cancel", SinkArgumentType.Bool)]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(64513)]
		void BeforeFolderSharingDialog([In, MarshalAs(UnmanagedType.IDispatch)] object folderToShare, [In] [Out] ref object cancel);
	}

}
