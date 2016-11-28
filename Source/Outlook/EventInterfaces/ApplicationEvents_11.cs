using System;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
using NetOffice;

namespace NetOffice.OutlookApi
{	
	#pragma warning disable
	
	#region SinkPoint Interface

	[SupportByVersionAttribute("Outlook", 11,12,14,15,16)]
	[ComImport, Guid("0006302C-0000-0000-C000-000000000046"), InterfaceType(ComInterfaceType.InterfaceIsIDispatch), TypeLibType((short)0x1010)]
	public interface ApplicationEvents_11
	{
		[SupportByVersionAttribute("Outlook", 11,12,14,15,16)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(61442)]
		void ItemSend([In, MarshalAs(UnmanagedType.IDispatch)] object item, [In] [Out] ref object cancel);

		[SupportByVersionAttribute("Outlook", 11,12,14,15,16)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(61443)]
		void NewMail();

		[SupportByVersionAttribute("Outlook", 11,12,14,15,16)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(61444)]
		void Reminder([In, MarshalAs(UnmanagedType.IDispatch)] object item);

		[SupportByVersionAttribute("Outlook", 11,12,14,15,16)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(61445)]
		void OptionsPagesAdd([In, MarshalAs(UnmanagedType.IDispatch)] object pages);

		[SupportByVersionAttribute("Outlook", 11,12,14,15,16)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(61446)]
		void Startup();

		[SupportByVersionAttribute("Outlook", 11,12,14,15,16)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(61447)]
		void Quit();

		[SupportByVersionAttribute("Outlook", 11,12,14,15,16)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(64106)]
		void AdvancedSearchComplete([In, MarshalAs(UnmanagedType.IDispatch)] object searchObject);

		[SupportByVersionAttribute("Outlook", 11,12,14,15,16)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(64107)]
		void AdvancedSearchStopped([In, MarshalAs(UnmanagedType.IDispatch)] object searchObject);

		[SupportByVersionAttribute("Outlook", 11,12,14,15,16)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(64144)]
		void MAPILogonComplete();

		[SupportByVersionAttribute("Outlook", 11,12,14,15,16)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(64181)]
		void NewMailEx([In] object entryIDCollection);

		[SupportByVersionAttribute("Outlook", 12,14,15,16)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(64318)]
		void AttachmentContextMenuDisplay([In, MarshalAs(UnmanagedType.IDispatch)] object commandBar, [In, MarshalAs(UnmanagedType.IDispatch)] object attachments);

		[SupportByVersionAttribute("Outlook", 12,14,15,16)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(64322)]
		void FolderContextMenuDisplay([In, MarshalAs(UnmanagedType.IDispatch)] object commandBar, [In, MarshalAs(UnmanagedType.IDispatch)] object folder);

		[SupportByVersionAttribute("Outlook", 12,14,15,16)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(64323)]
		void StoreContextMenuDisplay([In, MarshalAs(UnmanagedType.IDispatch)] object commandBar, [In, MarshalAs(UnmanagedType.IDispatch)] object store);

		[SupportByVersionAttribute("Outlook", 12,14,15,16)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(64324)]
		void ShortcutContextMenuDisplay([In, MarshalAs(UnmanagedType.IDispatch)] object commandBar, [In, MarshalAs(UnmanagedType.IDispatch)] object shortcut);

		[SupportByVersionAttribute("Outlook", 12,14,15,16)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(64320)]
		void ViewContextMenuDisplay([In, MarshalAs(UnmanagedType.IDispatch)] object commandBar, [In, MarshalAs(UnmanagedType.IDispatch)] object view);

		[SupportByVersionAttribute("Outlook", 12,14,15,16)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(64321)]
		void ItemContextMenuDisplay([In, MarshalAs(UnmanagedType.IDispatch)] object commandBar, [In, MarshalAs(UnmanagedType.IDispatch)] object selection);

		[SupportByVersionAttribute("Outlook", 12,14,15,16)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(64422)]
		void ContextMenuClose([In] object contextMenu);

		[SupportByVersionAttribute("Outlook", 12,14,15,16)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(64423)]
		void ItemLoad([In, MarshalAs(UnmanagedType.IDispatch)] object item);

		[SupportByVersionAttribute("Outlook", 12,14,15,16)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(64513)]
		void BeforeFolderSharingDialog([In, MarshalAs(UnmanagedType.IDispatch)] object folderToShare, [In] [Out] ref object cancel);
	}
	
	#endregion
	
	#region SinkHelper
	
	[ComVisible(true), ClassInterface(ClassInterfaceType.None), TypeLibType(TypeLibTypeFlags.FHidden)]
	public class ApplicationEvents_11_SinkHelper : SinkHelper, ApplicationEvents_11
	{
		#region Static
		
		public static readonly string Id = "0006302C-0000-0000-C000-000000000046";
		
		#endregion
	
		#region Fields

		private IEventBinding	_eventBinding;
        private COMObject		_eventClass;
        
		#endregion
		
		#region Construction

		public ApplicationEvents_11_SinkHelper(COMObject eventClass, IConnectionPoint connectPoint): base(eventClass)
		{
			_eventClass = eventClass;
			_eventBinding = (IEventBinding)eventClass;
			SetupEventBinding(connectPoint);
		}
		
		#endregion
		
		#region Properties

        internal Core Factory
        {
            get
            {
                if (null != _eventClass)
                    return _eventClass.Factory;
                else
                    return Core.Default;
            }
        }

        #endregion

		#region ApplicationEvents_11 Members
		
		public void ItemSend([In, MarshalAs(UnmanagedType.IDispatch)] object item, [In] [Out] ref object cancel)
		{
			Delegate[] recipients = _eventBinding.GetEventRecipients("ItemSend");
			if( (true == _eventClass.IsCurrentlyDisposing) || (recipients.Length == 0) )
			{
				Invoker.ReleaseParamsArray(item, cancel);
				return;
			}

			object newItem = Factory.CreateObjectFromComProxy(_eventClass, item) as object;
			object[] paramsArray = new object[2];
			paramsArray[0] = newItem;
			paramsArray.SetValue(cancel, 1);
			_eventBinding.RaiseCustomEvent("ItemSend", ref paramsArray);

			cancel = (bool)paramsArray[1];
		}

		public void NewMail()
		{
			Delegate[] recipients = _eventBinding.GetEventRecipients("NewMail");
			if( (true == _eventClass.IsCurrentlyDisposing) || (recipients.Length == 0) )
			{
				Invoker.ReleaseParamsArray();
				return;
			}

			object[] paramsArray = new object[0];
			_eventBinding.RaiseCustomEvent("NewMail", ref paramsArray);
		}

		public void Reminder([In, MarshalAs(UnmanagedType.IDispatch)] object item)
		{
			Delegate[] recipients = _eventBinding.GetEventRecipients("Reminder");
			if( (true == _eventClass.IsCurrentlyDisposing) || (recipients.Length == 0) )
			{
				Invoker.ReleaseParamsArray(item);
				return;
			}

			object newItem = Factory.CreateObjectFromComProxy(_eventClass, item) as object;
			object[] paramsArray = new object[1];
			paramsArray[0] = newItem;
			_eventBinding.RaiseCustomEvent("Reminder", ref paramsArray);
		}

		public void OptionsPagesAdd([In, MarshalAs(UnmanagedType.IDispatch)] object pages)
		{
			Delegate[] recipients = _eventBinding.GetEventRecipients("OptionsPagesAdd");
			if( (true == _eventClass.IsCurrentlyDisposing) || (recipients.Length == 0) )
			{
				Invoker.ReleaseParamsArray(pages);
				return;
			}

			NetOffice.OutlookApi.PropertyPages newPages = Factory.CreateObjectFromComProxy(_eventClass, pages) as NetOffice.OutlookApi.PropertyPages;
			object[] paramsArray = new object[1];
			paramsArray[0] = newPages;
			_eventBinding.RaiseCustomEvent("OptionsPagesAdd", ref paramsArray);
		}

		public void Startup()
		{
			Delegate[] recipients = _eventBinding.GetEventRecipients("Startup");
			if( (true == _eventClass.IsCurrentlyDisposing) || (recipients.Length == 0) )
			{
				Invoker.ReleaseParamsArray();
				return;
			}

			object[] paramsArray = new object[0];
			_eventBinding.RaiseCustomEvent("Startup", ref paramsArray);
		}

		public void Quit()
		{
			Delegate[] recipients = _eventBinding.GetEventRecipients("Quit");
			if( (true == _eventClass.IsCurrentlyDisposing) || (recipients.Length == 0) )
			{
				Invoker.ReleaseParamsArray();
				return;
			}

			object[] paramsArray = new object[0];
			_eventBinding.RaiseCustomEvent("Quit", ref paramsArray);
		}

		public void AdvancedSearchComplete([In, MarshalAs(UnmanagedType.IDispatch)] object searchObject)
		{
			Delegate[] recipients = _eventBinding.GetEventRecipients("AdvancedSearchComplete");
			if( (true == _eventClass.IsCurrentlyDisposing) || (recipients.Length == 0) )
			{
				Invoker.ReleaseParamsArray(searchObject);
				return;
			}

			NetOffice.OutlookApi.Search newSearchObject = Factory.CreateObjectFromComProxy(_eventClass, searchObject) as NetOffice.OutlookApi.Search;
			object[] paramsArray = new object[1];
			paramsArray[0] = newSearchObject;
			_eventBinding.RaiseCustomEvent("AdvancedSearchComplete", ref paramsArray);
		}

		public void AdvancedSearchStopped([In, MarshalAs(UnmanagedType.IDispatch)] object searchObject)
		{
			Delegate[] recipients = _eventBinding.GetEventRecipients("AdvancedSearchStopped");
			if( (true == _eventClass.IsCurrentlyDisposing) || (recipients.Length == 0) )
			{
				Invoker.ReleaseParamsArray(searchObject);
				return;
			}

			NetOffice.OutlookApi.Search newSearchObject = Factory.CreateObjectFromComProxy(_eventClass, searchObject) as NetOffice.OutlookApi.Search;
			object[] paramsArray = new object[1];
			paramsArray[0] = newSearchObject;
			_eventBinding.RaiseCustomEvent("AdvancedSearchStopped", ref paramsArray);
		}

		public void MAPILogonComplete()
		{
			Delegate[] recipients = _eventBinding.GetEventRecipients("MAPILogonComplete");
			if( (true == _eventClass.IsCurrentlyDisposing) || (recipients.Length == 0) )
			{
				Invoker.ReleaseParamsArray();
				return;
			}

			object[] paramsArray = new object[0];
			_eventBinding.RaiseCustomEvent("MAPILogonComplete", ref paramsArray);
		}

		public void NewMailEx([In] object entryIDCollection)
		{
			Delegate[] recipients = _eventBinding.GetEventRecipients("NewMailEx");
			if( (true == _eventClass.IsCurrentlyDisposing) || (recipients.Length == 0) )
			{
				Invoker.ReleaseParamsArray(entryIDCollection);
				return;
			}

			string newEntryIDCollection = Convert.ToString(entryIDCollection);
			object[] paramsArray = new object[1];
			paramsArray[0] = newEntryIDCollection;
			_eventBinding.RaiseCustomEvent("NewMailEx", ref paramsArray);
		}

		public void AttachmentContextMenuDisplay([In, MarshalAs(UnmanagedType.IDispatch)] object commandBar, [In, MarshalAs(UnmanagedType.IDispatch)] object attachments)
		{
			Delegate[] recipients = _eventBinding.GetEventRecipients("AttachmentContextMenuDisplay");
			if( (true == _eventClass.IsCurrentlyDisposing) || (recipients.Length == 0) )
			{
				Invoker.ReleaseParamsArray(commandBar, attachments);
				return;
			}

			NetOffice.OfficeApi.CommandBar newCommandBar = Factory.CreateObjectFromComProxy(_eventClass, commandBar) as NetOffice.OfficeApi.CommandBar;
			NetOffice.OutlookApi.AttachmentSelection newAttachments = Factory.CreateObjectFromComProxy(_eventClass, attachments) as NetOffice.OutlookApi.AttachmentSelection;
			object[] paramsArray = new object[2];
			paramsArray[0] = newCommandBar;
			paramsArray[1] = newAttachments;
			_eventBinding.RaiseCustomEvent("AttachmentContextMenuDisplay", ref paramsArray);
		}

		public void FolderContextMenuDisplay([In, MarshalAs(UnmanagedType.IDispatch)] object commandBar, [In, MarshalAs(UnmanagedType.IDispatch)] object folder)
		{
			Delegate[] recipients = _eventBinding.GetEventRecipients("FolderContextMenuDisplay");
			if( (true == _eventClass.IsCurrentlyDisposing) || (recipients.Length == 0) )
			{
				Invoker.ReleaseParamsArray(commandBar, folder);
				return;
			}

			NetOffice.OfficeApi.CommandBar newCommandBar = Factory.CreateObjectFromComProxy(_eventClass, commandBar) as NetOffice.OfficeApi.CommandBar;
			NetOffice.OutlookApi.Folder newFolder = Factory.CreateObjectFromComProxy(_eventClass, folder) as NetOffice.OutlookApi.Folder;
			object[] paramsArray = new object[2];
			paramsArray[0] = newCommandBar;
			paramsArray[1] = newFolder;
			_eventBinding.RaiseCustomEvent("FolderContextMenuDisplay", ref paramsArray);
		}

		public void StoreContextMenuDisplay([In, MarshalAs(UnmanagedType.IDispatch)] object commandBar, [In, MarshalAs(UnmanagedType.IDispatch)] object store)
		{
			Delegate[] recipients = _eventBinding.GetEventRecipients("StoreContextMenuDisplay");
			if( (true == _eventClass.IsCurrentlyDisposing) || (recipients.Length == 0) )
			{
				Invoker.ReleaseParamsArray(commandBar, store);
				return;
			}

			NetOffice.OfficeApi.CommandBar newCommandBar = Factory.CreateObjectFromComProxy(_eventClass, commandBar) as NetOffice.OfficeApi.CommandBar;
			NetOffice.OutlookApi.Store newStore = Factory.CreateObjectFromComProxy(_eventClass, store) as NetOffice.OutlookApi.Store;
			object[] paramsArray = new object[2];
			paramsArray[0] = newCommandBar;
			paramsArray[1] = newStore;
			_eventBinding.RaiseCustomEvent("StoreContextMenuDisplay", ref paramsArray);
		}

		public void ShortcutContextMenuDisplay([In, MarshalAs(UnmanagedType.IDispatch)] object commandBar, [In, MarshalAs(UnmanagedType.IDispatch)] object shortcut)
		{
			Delegate[] recipients = _eventBinding.GetEventRecipients("ShortcutContextMenuDisplay");
			if( (true == _eventClass.IsCurrentlyDisposing) || (recipients.Length == 0) )
			{
				Invoker.ReleaseParamsArray(commandBar, shortcut);
				return;
			}

			NetOffice.OfficeApi.CommandBar newCommandBar = Factory.CreateObjectFromComProxy(_eventClass, commandBar) as NetOffice.OfficeApi.CommandBar;
			NetOffice.OutlookApi.OutlookBarShortcut newShortcut = Factory.CreateObjectFromComProxy(_eventClass, shortcut) as NetOffice.OutlookApi.OutlookBarShortcut;
			object[] paramsArray = new object[2];
			paramsArray[0] = newCommandBar;
			paramsArray[1] = newShortcut;
			_eventBinding.RaiseCustomEvent("ShortcutContextMenuDisplay", ref paramsArray);
		}

		public void ViewContextMenuDisplay([In, MarshalAs(UnmanagedType.IDispatch)] object commandBar, [In, MarshalAs(UnmanagedType.IDispatch)] object view)
		{
			Delegate[] recipients = _eventBinding.GetEventRecipients("ViewContextMenuDisplay");
			if( (true == _eventClass.IsCurrentlyDisposing) || (recipients.Length == 0) )
			{
				Invoker.ReleaseParamsArray(commandBar, view);
				return;
			}

			NetOffice.OfficeApi.CommandBar newCommandBar = Factory.CreateObjectFromComProxy(_eventClass, commandBar) as NetOffice.OfficeApi.CommandBar;
			NetOffice.OutlookApi.View newView = Factory.CreateObjectFromComProxy(_eventClass, view) as NetOffice.OutlookApi.View;
			object[] paramsArray = new object[2];
			paramsArray[0] = newCommandBar;
			paramsArray[1] = newView;
			_eventBinding.RaiseCustomEvent("ViewContextMenuDisplay", ref paramsArray);
		}

		public void ItemContextMenuDisplay([In, MarshalAs(UnmanagedType.IDispatch)] object commandBar, [In, MarshalAs(UnmanagedType.IDispatch)] object selection)
		{
			Delegate[] recipients = _eventBinding.GetEventRecipients("ItemContextMenuDisplay");
			if( (true == _eventClass.IsCurrentlyDisposing) || (recipients.Length == 0) )
			{
				Invoker.ReleaseParamsArray(commandBar, selection);
				return;
			}

			NetOffice.OfficeApi.CommandBar newCommandBar = Factory.CreateObjectFromComProxy(_eventClass, commandBar) as NetOffice.OfficeApi.CommandBar;
			NetOffice.OutlookApi.Selection newSelection = Factory.CreateObjectFromComProxy(_eventClass, selection) as NetOffice.OutlookApi.Selection;
			object[] paramsArray = new object[2];
			paramsArray[0] = newCommandBar;
			paramsArray[1] = newSelection;
			_eventBinding.RaiseCustomEvent("ItemContextMenuDisplay", ref paramsArray);
		}

		public void ContextMenuClose([In] object contextMenu)
		{
			Delegate[] recipients = _eventBinding.GetEventRecipients("ContextMenuClose");
			if( (true == _eventClass.IsCurrentlyDisposing) || (recipients.Length == 0) )
			{
				Invoker.ReleaseParamsArray(contextMenu);
				return;
			}

			NetOffice.OutlookApi.Enums.OlContextMenu newContextMenu = (NetOffice.OutlookApi.Enums.OlContextMenu)contextMenu;
			object[] paramsArray = new object[1];
			paramsArray[0] = newContextMenu;
			_eventBinding.RaiseCustomEvent("ContextMenuClose", ref paramsArray);
		}

		public void ItemLoad([In, MarshalAs(UnmanagedType.IDispatch)] object item)
		{
			Delegate[] recipients = _eventBinding.GetEventRecipients("ItemLoad");
			if( (true == _eventClass.IsCurrentlyDisposing) || (recipients.Length == 0) )
			{
				Invoker.ReleaseParamsArray(item);
				return;
			}

			object newItem = Factory.CreateObjectFromComProxy(_eventClass, item) as object;
			object[] paramsArray = new object[1];
			paramsArray[0] = newItem;
			_eventBinding.RaiseCustomEvent("ItemLoad", ref paramsArray);
		}

		public void BeforeFolderSharingDialog([In, MarshalAs(UnmanagedType.IDispatch)] object folderToShare, [In] [Out] ref object cancel)
		{
			Delegate[] recipients = _eventBinding.GetEventRecipients("BeforeFolderSharingDialog");
			if( (true == _eventClass.IsCurrentlyDisposing) || (recipients.Length == 0) )
			{
				Invoker.ReleaseParamsArray(folderToShare, cancel);
				return;
			}

			NetOffice.OutlookApi.MAPIFolder newFolderToShare = Factory.CreateObjectFromComProxy(_eventClass, folderToShare) as NetOffice.OutlookApi.MAPIFolder;
			object[] paramsArray = new object[2];
			paramsArray[0] = newFolderToShare;
			paramsArray.SetValue(cancel, 1);
			_eventBinding.RaiseCustomEvent("BeforeFolderSharingDialog", ref paramsArray);

			cancel = (bool)paramsArray[1];
		}

		#endregion
	}
	
	#endregion
	
	#pragma warning restore
}