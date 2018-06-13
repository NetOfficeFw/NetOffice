using System;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
using NetOffice;
using NetOffice.Attributes;

namespace NetOffice.OutlookApi.Behind.EventContracts
{
    /// <summary>
    /// Default implementation of <see cref="NetOffice.OutlookApi.EventContracts.ApplicationEvents_11"/>
    /// </summary>
    [InternalEntity(InternalEntityKind.SinkHelper)]
    [ComVisible(true), ClassInterface(ClassInterfaceType.None), TypeLibType(TypeLibTypeFlags.FHidden)]
	public class ApplicationEvents_11_SinkHelper : SinkHelper, NetOffice.OutlookApi.EventContracts.ApplicationEvents_11
	{
        #region Static

        /// <summary>
        /// Interface Id from ApplicationEvents_11
        /// </summary>
        public static readonly string Id = "0006302C-0000-0000-C000-000000000046";
		
		#endregion
	
		#region Ctor

        /// <summary>
        /// Creates an instance of the class
        /// </summary>
        /// <param name="eventClass"></param>
        /// <param name="connectPoint"></param>
		public ApplicationEvents_11_SinkHelper(ICOMObject eventClass, IConnectionPoint connectPoint): base(eventClass)
		{
			SetupEventBinding(connectPoint);		}
		
		#endregion

		#region ApplicationEvents_11
		
        /// <summary>
        /// 
        /// </summary>
        /// <param name="item"></param>
        /// <param name="cancel"></param>
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

        /// <summary>
        /// 
        /// </summary>
		public void NewMail()
        {
            if (!Validate("NewMail"))
            {
                return;
            }

			object[] paramsArray = new object[0];
			EventBinding.RaiseCustomEvent("NewMail", ref paramsArray);
		}

        /// <summary>
        /// 
        /// </summary>
        /// <param name="item"></param>
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

        /// <summary>
        /// 
        /// </summary>
        /// <param name="pages"></param>
		public void OptionsPagesAdd([In, MarshalAs(UnmanagedType.IDispatch)] object pages)
		{
            if (!Validate("OptionsPagesAdd"))
            {
                Invoker.ReleaseParamsArray(pages);
                return;
            }
			NetOffice.OutlookApi.PropertyPages newPages = Factory.CreateKnownObjectFromComProxy<NetOffice.OutlookApi.PropertyPages>(EventClass, pages, typeof(NetOffice.OutlookApi.PropertyPages));
			object[] paramsArray = new object[1];
			paramsArray[0] = newPages;
			EventBinding.RaiseCustomEvent("OptionsPagesAdd", ref paramsArray);
		}

        /// <summary>
        /// 
        /// </summary>
		public void Startup()
		{
            if (!Validate("Startup"))
            {
                return;
            }

			object[] paramsArray = new object[0];
			EventBinding.RaiseCustomEvent("Startup", ref paramsArray);
		}

        /// <summary>
        /// 
        /// </summary>
		public void Quit()
        {
            if (!Validate("Startup"))
            {
                return;
            }

			object[] paramsArray = new object[0];
			EventBinding.RaiseCustomEvent("Quit", ref paramsArray);
		}

        /// <summary>
        /// 
        /// </summary>
        /// <param name="searchObject"></param>
		public void AdvancedSearchComplete([In, MarshalAs(UnmanagedType.IDispatch)] object searchObject)
		{
            if (!Validate("AdvancedSearchComplete"))
            {
                Invoker.ReleaseParamsArray(searchObject);
                return;
            }

			NetOffice.OutlookApi.Search newSearchObject = Factory.CreateKnownObjectFromComProxy<NetOffice.OutlookApi.Search>(EventClass, searchObject, typeof(NetOffice.OutlookApi.Search));
			object[] paramsArray = new object[1];
			paramsArray[0] = newSearchObject;
			EventBinding.RaiseCustomEvent("AdvancedSearchComplete", ref paramsArray);
		}

        /// <summary>
        /// 
        /// </summary>
        /// <param name="searchObject"></param>
		public void AdvancedSearchStopped([In, MarshalAs(UnmanagedType.IDispatch)] object searchObject)
		{
            if (!Validate("AdvancedSearchStopped"))
            {
                Invoker.ReleaseParamsArray(searchObject);
                return;
            }

			NetOffice.OutlookApi.Search newSearchObject = Factory.CreateKnownObjectFromComProxy<NetOffice.OutlookApi.Search>(EventClass, searchObject, typeof(NetOffice.OutlookApi.Search));
			object[] paramsArray = new object[1];
			paramsArray[0] = newSearchObject;
			EventBinding.RaiseCustomEvent("AdvancedSearchStopped", ref paramsArray);
		}

        /// <summary>
        /// 
        /// </summary>
		public void MAPILogonComplete()
		{
            if (!Validate("MAPILogonComplete"))
            {
                return;
            }

            object[] paramsArray = new object[0];
			EventBinding.RaiseCustomEvent("MAPILogonComplete", ref paramsArray);
		}

        /// <summary>
        /// 
        /// </summary>
        /// <param name="entryIDCollection"></param>
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

        /// <summary>
        /// 
        /// </summary>
        /// <param name="commandBar"></param>
        /// <param name="attachments"></param>
        public void AttachmentContextMenuDisplay([In, MarshalAs(UnmanagedType.IDispatch)] object commandBar, [In, MarshalAs(UnmanagedType.IDispatch)] object attachments)
		{
            if (!Validate("AttachmentContextMenuDisplay"))
            {
                Invoker.ReleaseParamsArray(commandBar, attachments);
                return;
            }

			NetOffice.OfficeApi.CommandBar newCommandBar = Factory.CreateKnownObjectFromComProxy<NetOffice.OfficeApi.CommandBar>(EventClass, commandBar, typeof(NetOffice.OfficeApi.CommandBar));
			NetOffice.OutlookApi.AttachmentSelection newAttachments = Factory.CreateKnownObjectFromComProxy<NetOffice.OutlookApi.AttachmentSelection>(EventClass, attachments, typeof(NetOffice.OutlookApi.AttachmentSelection));
			object[] paramsArray = new object[2];
			paramsArray[0] = newCommandBar;
			paramsArray[1] = newAttachments;
			EventBinding.RaiseCustomEvent("AttachmentContextMenuDisplay", ref paramsArray);
		}

        /// <summary>
        /// 
        /// </summary>
        /// <param name="commandBar"></param>
        /// <param name="folder"></param>
        public void FolderContextMenuDisplay([In, MarshalAs(UnmanagedType.IDispatch)] object commandBar, [In, MarshalAs(UnmanagedType.IDispatch)] object folder)
		{
            if (!Validate("FolderContextMenuDisplay"))
            {
                Invoker.ReleaseParamsArray(commandBar, folder);
                return;
            }

            NetOffice.OfficeApi.CommandBar newCommandBar = Factory.CreateKnownObjectFromComProxy<NetOffice.OfficeApi.CommandBar>(EventClass, commandBar, typeof(NetOffice.OfficeApi.CommandBar));
            NetOffice.OutlookApi.Folder newFolder = Factory.CreateKnownObjectFromComProxy<NetOffice.OutlookApi.Folder>(EventClass, folder, typeof(NetOffice.OutlookApi.Folder));
			object[] paramsArray = new object[2];
			paramsArray[0] = newCommandBar;
			paramsArray[1] = newFolder;
			EventBinding.RaiseCustomEvent("FolderContextMenuDisplay", ref paramsArray);
		}

        /// <summary>
        /// 
        /// </summary>
        /// <param name="commandBar"></param>
        /// <param name="store"></param>
        public void StoreContextMenuDisplay([In, MarshalAs(UnmanagedType.IDispatch)] object commandBar, [In, MarshalAs(UnmanagedType.IDispatch)] object store)
		{
            if (!Validate("StoreContextMenuDisplay"))
            {
                Invoker.ReleaseParamsArray(commandBar, store);
                return;
            }

            NetOffice.OfficeApi.CommandBar newCommandBar = Factory.CreateKnownObjectFromComProxy<NetOffice.OfficeApi.CommandBar>(EventClass, commandBar, typeof(NetOffice.OfficeApi.CommandBar));
            NetOffice.OutlookApi.Store newStore = Factory.CreateKnownObjectFromComProxy<NetOffice.OutlookApi.Store>(EventClass, store, typeof(NetOffice.OutlookApi.Store));
			object[] paramsArray = new object[2];
			paramsArray[0] = newCommandBar;
			paramsArray[1] = newStore;
			EventBinding.RaiseCustomEvent("StoreContextMenuDisplay", ref paramsArray);
		}

        /// <summary>
        /// 
        /// </summary>
        /// <param name="commandBar"></param>
        /// <param name="shortcut"></param>
        public void ShortcutContextMenuDisplay([In, MarshalAs(UnmanagedType.IDispatch)] object commandBar, [In, MarshalAs(UnmanagedType.IDispatch)] object shortcut)
		{
            if (!Validate("ShortcutContextMenuDisplay"))
            {
                Invoker.ReleaseParamsArray(commandBar, shortcut);
                return;
            }

            NetOffice.OfficeApi.CommandBar newCommandBar = Factory.CreateKnownObjectFromComProxy<NetOffice.OfficeApi.CommandBar>(EventClass, commandBar, typeof(NetOffice.OfficeApi.CommandBar));
            NetOffice.OutlookApi.OutlookBarShortcut newShortcut = Factory.CreateKnownObjectFromComProxy<NetOffice.OutlookApi.OutlookBarShortcut>(EventClass, shortcut, typeof(NetOffice.OutlookApi.OutlookBarShortcut));
			object[] paramsArray = new object[2];
			paramsArray[0] = newCommandBar;
			paramsArray[1] = newShortcut;
			EventBinding.RaiseCustomEvent("ShortcutContextMenuDisplay", ref paramsArray);
		}

        /// <summary>
        /// 
        /// </summary>
        /// <param name="commandBar"></param>
        /// <param name="view"></param>
        public void ViewContextMenuDisplay([In, MarshalAs(UnmanagedType.IDispatch)] object commandBar, [In, MarshalAs(UnmanagedType.IDispatch)] object view)
        {
            if (!Validate("ViewContextMenuDisplay"))
            {
                Invoker.ReleaseParamsArray(commandBar, view);
                return;
            }

            NetOffice.OfficeApi.CommandBar newCommandBar = Factory.CreateKnownObjectFromComProxy<NetOffice.OfficeApi.CommandBar>(EventClass, commandBar, typeof(NetOffice.OfficeApi.CommandBar));
            NetOffice.OutlookApi.View newView = Factory.CreateKnownObjectFromComProxy<NetOffice.OutlookApi.View>(EventClass, view, typeof(NetOffice.OutlookApi.View));
			object[] paramsArray = new object[2];
			paramsArray[0] = newCommandBar;
			paramsArray[1] = newView;
			EventBinding.RaiseCustomEvent("ViewContextMenuDisplay", ref paramsArray);
		}

        /// <summary>
        /// 
        /// </summary>
        /// <param name="commandBar"></param>
        /// <param name="selection"></param>
        public void ItemContextMenuDisplay([In, MarshalAs(UnmanagedType.IDispatch)] object commandBar, [In, MarshalAs(UnmanagedType.IDispatch)] object selection)
		{
            if (!Validate("ItemContextMenuDisplay"))
            {
                Invoker.ReleaseParamsArray(commandBar, selection);
                return;
            }

            NetOffice.OfficeApi.CommandBar newCommandBar = Factory.CreateKnownObjectFromComProxy<NetOffice.OfficeApi.CommandBar>(EventClass, commandBar, typeof(NetOffice.OfficeApi.CommandBar));
            NetOffice.OutlookApi.Selection newSelection = Factory.CreateKnownObjectFromComProxy<NetOffice.OutlookApi.Selection>(EventClass, selection, typeof(NetOffice.OutlookApi.Selection));
			object[] paramsArray = new object[2];
			paramsArray[0] = newCommandBar;
			paramsArray[1] = newSelection;
			EventBinding.RaiseCustomEvent("ItemContextMenuDisplay", ref paramsArray);
		}

        /// <summary>
        /// 
        /// </summary>
        /// <param name="contextMenu"></param>
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

        /// <summary>
        /// 
        /// </summary>
        /// <param name="item"></param>
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

        /// <summary>
        /// 
        /// </summary>
        /// <param name="folderToShare"></param>
        /// <param name="cancel"></param>
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
}

