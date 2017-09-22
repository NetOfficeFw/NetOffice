using System;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
using NetOffice;
using NetOffice.Attributes;

namespace NetOffice.OutlookApi.Events
{	
	#pragma warning disable
	
	#region SinkPoint Interface

	[SupportByVersion("Outlook", 11,12,14,15,16)]
    [InternalEntity(InternalEntityKind.ComEventInterface)]
    [ComImport, Guid("0006302C-0000-0000-C000-000000000046"), InterfaceType(ComInterfaceType.InterfaceIsIDispatch), TypeLibType((short)0x1010)]
	public interface ApplicationEvents_11
	{
		[SupportByVersion("Outlook", 11,12,14,15,16)]
        [SinkArgument("item", SinkArgumentType.UnknownProxy)]
        [SinkArgument("cancel", SinkArgumentType.Bool)]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(61442)]
		void ItemSend([In, MarshalAs(UnmanagedType.IDispatch)] object item, [In] [Out] ref object cancel);

		[SupportByVersion("Outlook", 11,12,14,15,16)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(61443)]
		void NewMail();

		[SupportByVersion("Outlook", 11,12,14,15,16)]
        [SinkArgument("item", SinkArgumentType.UnknownProxy)]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(61444)]
		void Reminder([In, MarshalAs(UnmanagedType.IDispatch)] object item);

		[SupportByVersion("Outlook", 11,12,14,15,16)]
        [SinkArgument("pages", typeof(OutlookApi.PropertyPages))]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(61445)]
		void OptionsPagesAdd([In, MarshalAs(UnmanagedType.IDispatch)] object pages);

		[SupportByVersion("Outlook", 11,12,14,15,16)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(61446)]
		void Startup();

		[SupportByVersion("Outlook", 11,12,14,15,16)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(61447)]
		void Quit();

		[SupportByVersion("Outlook", 11,12,14,15,16)]
        [SinkArgument("searchObject", typeof(OutlookApi.Search))]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(64106)]
		void AdvancedSearchComplete([In, MarshalAs(UnmanagedType.IDispatch)] object searchObject);

		[SupportByVersion("Outlook", 11,12,14,15,16)]
        [SinkArgument("searchObject", typeof(OutlookApi.Search))]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(64107)]
		void AdvancedSearchStopped([In, MarshalAs(UnmanagedType.IDispatch)] object searchObject);

		[SupportByVersion("Outlook", 11,12,14,15,16)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(64144)]
		void MAPILogonComplete();

		[SupportByVersion("Outlook", 11,12,14,15,16)]
        [SinkArgument("entryIDCollection", SinkArgumentType.String)]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(64181)]
		void NewMailEx([In] object entryIDCollection);

		[SupportByVersion("Outlook", 12,14,15,16)]
        [SinkArgument("commandBar", typeof(OfficeApi.CommandBar))]
        [SinkArgument("attachments", typeof(OutlookApi.AttachmentSelection))]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(64318)]
		void AttachmentContextMenuDisplay([In, MarshalAs(UnmanagedType.IDispatch)] object commandBar, [In, MarshalAs(UnmanagedType.IDispatch)] object attachments);

		[SupportByVersion("Outlook", 12,14,15,16)]
        [SinkArgument("commandBar", typeof(OfficeApi.CommandBar))]
        [SinkArgument("folder", typeof(OutlookApi.Folder))]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(64322)]
		void FolderContextMenuDisplay([In, MarshalAs(UnmanagedType.IDispatch)] object commandBar, [In, MarshalAs(UnmanagedType.IDispatch)] object folder);

		[SupportByVersion("Outlook", 12,14,15,16)]
        [SinkArgument("commandBar", typeof(OfficeApi.CommandBar))]
        [SinkArgument("store", typeof(OutlookApi.Store))]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(64323)]
		void StoreContextMenuDisplay([In, MarshalAs(UnmanagedType.IDispatch)] object commandBar, [In, MarshalAs(UnmanagedType.IDispatch)] object store);

		[SupportByVersion("Outlook", 12,14,15,16)]
        [SinkArgument("commandBar", typeof(OfficeApi.CommandBar))]
        [SinkArgument("shortcut", typeof(OutlookApi.OutlookBarShortcut))]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(64324)]
		void ShortcutContextMenuDisplay([In, MarshalAs(UnmanagedType.IDispatch)] object commandBar, [In, MarshalAs(UnmanagedType.IDispatch)] object shortcut);

		[SupportByVersion("Outlook", 12,14,15,16)]
        [SinkArgument("commandBar", typeof(OfficeApi.CommandBar))]
        [SinkArgument("view", typeof(OutlookApi.View))]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(64320)]
		void ViewContextMenuDisplay([In, MarshalAs(UnmanagedType.IDispatch)] object commandBar, [In, MarshalAs(UnmanagedType.IDispatch)] object view);

		[SupportByVersion("Outlook", 12,14,15,16)]
        [SinkArgument("commandBar", typeof(OfficeApi.CommandBar))]
        [SinkArgument("selection", typeof(OutlookApi.Selection))]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(64321)]
		void ItemContextMenuDisplay([In, MarshalAs(UnmanagedType.IDispatch)] object commandBar, [In, MarshalAs(UnmanagedType.IDispatch)] object selection);

		[SupportByVersion("Outlook", 12,14,15,16)]
        [SinkArgument("contextMenu", SinkArgumentType.Enum, typeof(OutlookApi.Enums.OlContextMenu))]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(64422)]
		void ContextMenuClose([In] object contextMenu);

		[SupportByVersion("Outlook", 12,14,15,16)]
        [SinkArgument("item", SinkArgumentType.UnknownProxy)]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(64423)]
		void ItemLoad([In, MarshalAs(UnmanagedType.IDispatch)] object item);

		[SupportByVersion("Outlook", 12,14,15,16)]
        [SinkArgument("folderToShare", typeof(OutlookApi.MAPIFolder))]
        [SinkArgument("cancel", SinkArgumentType.Bool)]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(64513)]
		void BeforeFolderSharingDialog([In, MarshalAs(UnmanagedType.IDispatch)] object folderToShare, [In] [Out] ref object cancel);
	}

    #endregion

    #region SinkHelper

    [InternalEntity(InternalEntityKind.SinkHelper)]
    [ComVisible(true), ClassInterface(ClassInterfaceType.None), TypeLibType(TypeLibTypeFlags.FHidden)]
	public class ApplicationEvents_11_SinkHelper : SinkHelper, ApplicationEvents_11
	{
		#region Static
		
		public static readonly string Id = "0006302C-0000-0000-C000-000000000046";
		
		#endregion
	
		#region Ctor

		public ApplicationEvents_11_SinkHelper(ICOMObject eventClass, IConnectionPoint connectPoint): base(eventClass)
		{
			SetupEventBinding(connectPoint);
		}
		
		#endregion

		#region ApplicationEvents_11
		
		public void ItemSend([In, MarshalAs(UnmanagedType.IDispatch)] object item, [In] [Out] ref object cancel)
		{
            if (!Validate("ItemSend"))
            {
                Invoker.ReleaseParamsArray(item, cancel);
                return;
            }

			object newItem = Factory.CreateEventArgumentObjectFromComProxy(EventClass, item) as object;
			object[] paramsArray = new object[2];
			paramsArray[0] = newItem;
			paramsArray.SetValue(cancel, 1);
			EventBinding.RaiseCustomEvent("ItemSend", ref paramsArray);

			cancel = ToBoolean(paramsArray[1]);
		}

		public void NewMail()
        {
            if (!Validate("NewMail"))
            {
                return;
            }

			object[] paramsArray = new object[0];
			EventBinding.RaiseCustomEvent("NewMail", ref paramsArray);
		}

		public void Reminder([In, MarshalAs(UnmanagedType.IDispatch)] object item)
		{
            if (!Validate("Reminder"))
            {
                Invoker.ReleaseParamsArray(item);
                return;
            }

			object newItem = Factory.CreateEventArgumentObjectFromComProxy(EventClass, item) as object;
			object[] paramsArray = new object[1];
			paramsArray[0] = newItem;
			EventBinding.RaiseCustomEvent("Reminder", ref paramsArray);
		}

		public void OptionsPagesAdd([In, MarshalAs(UnmanagedType.IDispatch)] object pages)
		{
            if (!Validate("OptionsPagesAdd"))
            {
                Invoker.ReleaseParamsArray(pages);
                return;
            }
			NetOffice.OutlookApi.PropertyPages newPages = Factory.CreateKnownObjectFromComProxy<NetOffice.OutlookApi.PropertyPages>(EventClass, pages, NetOffice.OutlookApi.PropertyPages.LateBindingApiWrapperType);
			object[] paramsArray = new object[1];
			paramsArray[0] = newPages;
			EventBinding.RaiseCustomEvent("OptionsPagesAdd", ref paramsArray);
		}

		public void Startup()
		{
            if (!Validate("Startup"))
            {
                return;
            }

			object[] paramsArray = new object[0];
			EventBinding.RaiseCustomEvent("Startup", ref paramsArray);
		}

		public void Quit()
        {
            if (!Validate("Startup"))
            {
                return;
            }

			object[] paramsArray = new object[0];
			EventBinding.RaiseCustomEvent("Quit", ref paramsArray);
		}

		public void AdvancedSearchComplete([In, MarshalAs(UnmanagedType.IDispatch)] object searchObject)
		{
            if (!Validate("AdvancedSearchComplete"))
            {
                Invoker.ReleaseParamsArray(searchObject);
                return;
            }

			NetOffice.OutlookApi.Search newSearchObject = Factory.CreateKnownObjectFromComProxy<NetOffice.OutlookApi.Search>(EventClass, searchObject, NetOffice.OutlookApi.Search.LateBindingApiWrapperType);
			object[] paramsArray = new object[1];
			paramsArray[0] = newSearchObject;
			EventBinding.RaiseCustomEvent("AdvancedSearchComplete", ref paramsArray);
		}

		public void AdvancedSearchStopped([In, MarshalAs(UnmanagedType.IDispatch)] object searchObject)
		{
            if (!Validate("AdvancedSearchStopped"))
            {
                Invoker.ReleaseParamsArray(searchObject);
                return;
            }

			NetOffice.OutlookApi.Search newSearchObject = Factory.CreateKnownObjectFromComProxy<NetOffice.OutlookApi.Search>(EventClass, searchObject, NetOffice.OutlookApi.Search.LateBindingApiWrapperType);
			object[] paramsArray = new object[1];
			paramsArray[0] = newSearchObject;
			EventBinding.RaiseCustomEvent("AdvancedSearchStopped", ref paramsArray);
		}

		public void MAPILogonComplete()
		{
            if (!Validate("MAPILogonComplete"))
            {
                return;
            }

            object[] paramsArray = new object[0];
			EventBinding.RaiseCustomEvent("MAPILogonComplete", ref paramsArray);
		}

		public void NewMailEx([In] object entryIDCollection)
        {
            if (!Validate("NewMailEx"))
            {
                Invoker.ReleaseParamsArray(entryIDCollection);
                return;
            }

			string newEntryIDCollection = ToString(entryIDCollection);
			object[] paramsArray = new object[1];
			paramsArray[0] = newEntryIDCollection;
			EventBinding.RaiseCustomEvent("NewMailEx", ref paramsArray);
		}

        public void AttachmentContextMenuDisplay([In, MarshalAs(UnmanagedType.IDispatch)] object commandBar, [In, MarshalAs(UnmanagedType.IDispatch)] object attachments)
		{
            if (!Validate("AttachmentContextMenuDisplay"))
            {
                Invoker.ReleaseParamsArray(commandBar, attachments);
                return;
            }

			NetOffice.OfficeApi.CommandBar newCommandBar = Factory.CreateKnownObjectFromComProxy<NetOffice.OfficeApi.CommandBar>(EventClass, commandBar, NetOffice.OfficeApi.CommandBar.LateBindingApiWrapperType);
			NetOffice.OutlookApi.AttachmentSelection newAttachments = Factory.CreateKnownObjectFromComProxy<NetOffice.OutlookApi.AttachmentSelection>(EventClass, attachments, NetOffice.OutlookApi.AttachmentSelection.LateBindingApiWrapperType);
			object[] paramsArray = new object[2];
			paramsArray[0] = newCommandBar;
			paramsArray[1] = newAttachments;
			EventBinding.RaiseCustomEvent("AttachmentContextMenuDisplay", ref paramsArray);
		}

        public void FolderContextMenuDisplay([In, MarshalAs(UnmanagedType.IDispatch)] object commandBar, [In, MarshalAs(UnmanagedType.IDispatch)] object folder)
		{
            if (!Validate("FolderContextMenuDisplay"))
            {
                Invoker.ReleaseParamsArray(commandBar, folder);
                return;
            }

            NetOffice.OfficeApi.CommandBar newCommandBar = Factory.CreateKnownObjectFromComProxy<NetOffice.OfficeApi.CommandBar>(EventClass, commandBar, NetOffice.OfficeApi.CommandBar.LateBindingApiWrapperType);
            NetOffice.OutlookApi.Folder newFolder = Factory.CreateKnownObjectFromComProxy<NetOffice.OutlookApi.Folder>(EventClass, folder, NetOffice.OutlookApi.Folder.LateBindingApiWrapperType);
			object[] paramsArray = new object[2];
			paramsArray[0] = newCommandBar;
			paramsArray[1] = newFolder;
			EventBinding.RaiseCustomEvent("FolderContextMenuDisplay", ref paramsArray);
		}

        public void StoreContextMenuDisplay([In, MarshalAs(UnmanagedType.IDispatch)] object commandBar, [In, MarshalAs(UnmanagedType.IDispatch)] object store)
		{
            if (!Validate("StoreContextMenuDisplay"))
            {
                Invoker.ReleaseParamsArray(commandBar, store);
                return;
            }

            NetOffice.OfficeApi.CommandBar newCommandBar = Factory.CreateKnownObjectFromComProxy<NetOffice.OfficeApi.CommandBar>(EventClass, commandBar, NetOffice.OfficeApi.CommandBar.LateBindingApiWrapperType);
            NetOffice.OutlookApi.Store newStore = Factory.CreateKnownObjectFromComProxy<NetOffice.OutlookApi.Store>(EventClass, store, NetOffice.OutlookApi.Store.LateBindingApiWrapperType);
			object[] paramsArray = new object[2];
			paramsArray[0] = newCommandBar;
			paramsArray[1] = newStore;
			EventBinding.RaiseCustomEvent("StoreContextMenuDisplay", ref paramsArray);
		}

        public void ShortcutContextMenuDisplay([In, MarshalAs(UnmanagedType.IDispatch)] object commandBar, [In, MarshalAs(UnmanagedType.IDispatch)] object shortcut)
		{
            if (!Validate("ShortcutContextMenuDisplay"))
            {
                Invoker.ReleaseParamsArray(commandBar, shortcut);
                return;
            }

            NetOffice.OfficeApi.CommandBar newCommandBar = Factory.CreateKnownObjectFromComProxy<NetOffice.OfficeApi.CommandBar>(EventClass, commandBar, NetOffice.OfficeApi.CommandBar.LateBindingApiWrapperType);
            NetOffice.OutlookApi.OutlookBarShortcut newShortcut = Factory.CreateKnownObjectFromComProxy<NetOffice.OutlookApi.OutlookBarShortcut>(EventClass, shortcut, NetOffice.OutlookApi.OutlookBarShortcut.LateBindingApiWrapperType);
			object[] paramsArray = new object[2];
			paramsArray[0] = newCommandBar;
			paramsArray[1] = newShortcut;
			EventBinding.RaiseCustomEvent("ShortcutContextMenuDisplay", ref paramsArray);
		}

        public void ViewContextMenuDisplay([In, MarshalAs(UnmanagedType.IDispatch)] object commandBar, [In, MarshalAs(UnmanagedType.IDispatch)] object view)
        {
            if (!Validate("ViewContextMenuDisplay"))
            {
                Invoker.ReleaseParamsArray(commandBar, view);
                return;
            }

            NetOffice.OfficeApi.CommandBar newCommandBar = Factory.CreateKnownObjectFromComProxy<NetOffice.OfficeApi.CommandBar>(EventClass, commandBar, NetOffice.OfficeApi.CommandBar.LateBindingApiWrapperType);
            NetOffice.OutlookApi.View newView = Factory.CreateKnownObjectFromComProxy<NetOffice.OutlookApi.View>(EventClass, view, NetOffice.OutlookApi.View.LateBindingApiWrapperType);
			object[] paramsArray = new object[2];
			paramsArray[0] = newCommandBar;
			paramsArray[1] = newView;
			EventBinding.RaiseCustomEvent("ViewContextMenuDisplay", ref paramsArray);
		}

        public void ItemContextMenuDisplay([In, MarshalAs(UnmanagedType.IDispatch)] object commandBar, [In, MarshalAs(UnmanagedType.IDispatch)] object selection)
		{
            if (!Validate("ItemContextMenuDisplay"))
            {
                Invoker.ReleaseParamsArray(commandBar, selection);
                return;
            }

            NetOffice.OfficeApi.CommandBar newCommandBar = Factory.CreateKnownObjectFromComProxy<NetOffice.OfficeApi.CommandBar>(EventClass, commandBar, NetOffice.OfficeApi.CommandBar.LateBindingApiWrapperType);
            NetOffice.OutlookApi.Selection newSelection = Factory.CreateKnownObjectFromComProxy<NetOffice.OutlookApi.Selection>(EventClass, selection, NetOffice.OutlookApi.Selection.LateBindingApiWrapperType);
			object[] paramsArray = new object[2];
			paramsArray[0] = newCommandBar;
			paramsArray[1] = newSelection;
			EventBinding.RaiseCustomEvent("ItemContextMenuDisplay", ref paramsArray);
		}

        public void ContextMenuClose([In] object contextMenu)
		{
            if (!Validate("ContextMenuClose"))
            {
                Invoker.ReleaseParamsArray(contextMenu);
                return;
            }

			NetOffice.OutlookApi.Enums.OlContextMenu newContextMenu = (NetOffice.OutlookApi.Enums.OlContextMenu)contextMenu;
			object[] paramsArray = new object[1];
			paramsArray[0] = newContextMenu;
			EventBinding.RaiseCustomEvent("ContextMenuClose", ref paramsArray);
		}

        public void ItemLoad([In, MarshalAs(UnmanagedType.IDispatch)] object item)
        {
            if (!Validate("ItemLoad"))
            {
                Invoker.ReleaseParamsArray(item);
                return;
            }

			object newItem = Factory.CreateEventArgumentObjectFromComProxy(EventClass, item) as object;
			object[] paramsArray = new object[1];
			paramsArray[0] = newItem;
			EventBinding.RaiseCustomEvent("ItemLoad", ref paramsArray);
		}

        public void BeforeFolderSharingDialog([In, MarshalAs(UnmanagedType.IDispatch)] object folderToShare, [In] [Out] ref object cancel)
		{
            if (!Validate("BeforeFolderSharingDialog"))
            {
                Invoker.ReleaseParamsArray(folderToShare, cancel);
                return;
            }

            NetOffice.OutlookApi.MAPIFolder newFolderToShare = Factory.CreateEventArgumentObjectFromComProxy(EventClass, folderToShare) as NetOffice.OutlookApi.MAPIFolder;
            object[] paramsArray = new object[2];
			paramsArray[0] = newFolderToShare;
			paramsArray.SetValue(cancel, 1);
			EventBinding.RaiseCustomEvent("BeforeFolderSharingDialog", ref paramsArray);

			cancel = ToBoolean(paramsArray[1]);
		}

		#endregion
	}
	
	#endregion
	
	#pragma warning restore
}