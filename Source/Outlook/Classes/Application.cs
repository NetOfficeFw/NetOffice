using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;
using NetOffice.Contribution.CollectionsGeneric;

namespace NetOffice.OutlookApi
{
	#region Delegates

	#pragma warning disable
	public delegate void Application_ItemSendEventHandler(ICOMObject item, ref bool cancel);
	public delegate void Application_NewMailEventHandler();
	public delegate void Application_ReminderEventHandler(ICOMObject item);
	public delegate void Application_OptionsPagesAddEventHandler(NetOffice.OutlookApi.PropertyPages pages);
	public delegate void Application_StartupEventHandler();
	public delegate void Application_QuitEventHandler();
	public delegate void Application_AdvancedSearchCompleteEventHandler(NetOffice.OutlookApi.Search searchObject);
	public delegate void Application_AdvancedSearchStoppedEventHandler(NetOffice.OutlookApi.Search searchObject);
	public delegate void Application_MAPILogonCompleteEventHandler();
	public delegate void Application_NewMailExEventHandler(string entryIDCollection);
	public delegate void Application_AttachmentContextMenuDisplayEventHandler(NetOffice.OfficeApi.CommandBar commandBar, NetOffice.OutlookApi.AttachmentSelection attachments);
	public delegate void Application_FolderContextMenuDisplayEventHandler(NetOffice.OfficeApi.CommandBar commandBar, NetOffice.OutlookApi.Folder folder);
	public delegate void Application_StoreContextMenuDisplayEventHandler(NetOffice.OfficeApi.CommandBar commandBar, NetOffice.OutlookApi.Store store);
	public delegate void Application_ShortcutContextMenuDisplayEventHandler(NetOffice.OfficeApi.CommandBar commandBar, NetOffice.OutlookApi.OutlookBarShortcut shortcut);
	public delegate void Application_ViewContextMenuDisplayEventHandler(NetOffice.OfficeApi.CommandBar commandBar, NetOffice.OutlookApi.View view);
	public delegate void Application_ItemContextMenuDisplayEventHandler(NetOffice.OfficeApi.CommandBar commandBar, NetOffice.OutlookApi.Selection selection);
	public delegate void Application_ContextMenuCloseEventHandler(NetOffice.OutlookApi.Enums.OlContextMenu contextMenu);
	public delegate void Application_ItemLoadEventHandler(ICOMObject item);
	public delegate void Application_BeforeFolderSharingDialogEventHandler(NetOffice.OutlookApi.MAPIFolder folderToShare, ref bool cancel);
	#pragma warning restore

	#endregion

	/// <summary>
	/// CoClass Application 
	/// SupportByVersion Outlook, 9,10,11,12,14,15,16
	/// </summary>
	/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff866895.aspx </remarks>
	[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
	[EntityType(EntityType.IsCoClass), ComProgId("Outlook.Application"), ModuleProvider(typeof(GlobalHelperModules.GlobalModule))]
	[EventSink(typeof(Events.ApplicationEvents_SinkHelper), typeof(Events.ApplicationEvents_10_SinkHelper), typeof(Events.ApplicationEvents_11_SinkHelper))]
	public class Application : _Application, IEventBinding
	{
		#pragma warning disable

		#region Fields
		
		private NetRuntimeSystem.Runtime.InteropServices.ComTypes.IConnectionPoint _connectPoint;
		private string _activeSinkId;
		private NetRuntimeSystem.Type _thisType;
		private Events.ApplicationEvents_SinkHelper _applicationEvents_SinkHelper;
        private Events.ApplicationEvents_10_SinkHelper _applicationEvents_10_SinkHelper;
        private Events.ApplicationEvents_11_SinkHelper _applicationEvents_11_SinkHelper;
	
		#endregion

		#region Type Information

        /// <summary>
        /// Instance Type
        /// </summary>
		[EditorBrowsable(EditorBrowsableState.Advanced), Browsable(false), Category("NetOffice"), CoreOverridden]
        public override Type InstanceType
        {
            get
            {
                return LateBindingApiWrapperType;
            }
        }

        private static Type _type;
		
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        public static Type LateBindingApiWrapperType
        {
            get
            {
                if (null == _type)
                    _type = typeof(Application);
                return _type;
            }
        }
        
        #endregion
        		
		#region Construction

		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
		public Application(Core factory, ICOMObject parentObject, object comProxy) : base(factory, parentObject, comProxy)
		{
			_callQuitInDispose = true;
			GlobalHelperModules.GlobalModule.Instance = this;
		}

        ///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
		public Application(ICOMObject parentObject, object comProxy) : base(parentObject, comProxy)
		{
			_callQuitInDispose = true;
			GlobalHelperModules.GlobalModule.Instance = this;
		}

		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Application(Core factory, ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(factory, parentObject, comProxy, comProxyType)
		{
			_callQuitInDispose = true;
		}

		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Application(ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(parentObject, comProxy, comProxyType)
		{
			_callQuitInDispose = true;
		}
		
		///<param name="replacedObject">object to replaced. replacedObject are not usable after this action</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Application(ICOMObject replacedObject) : base(replacedObject)
		{
			_callQuitInDispose = true;
		}
		
		/// <summary>
        /// Creates a new instance of Application 
        /// </summary>		
		public Application():base("Outlook.Application")
		{
			_callQuitInDispose = true;
			GlobalHelperModules.GlobalModule.Instance = this;
		}
		
		/// <summary>
        /// Creates a new instance of Application
        /// </summary>
        ///<param name="progId">registered ProgID</param>
		public Application(string progId):base(progId)
		{
			_callQuitInDispose = true;
			GlobalHelperModules.GlobalModule.Instance = this;
		}
		
        /// <summary>
		/// NetOffice method: dispose instance and all child instances
		/// </summary>
		/// <param name="disposeEventBinding">dispose event exported proxies with one or more event recipients</param>
		[Category("NetOffice"), CoreOverridden]
		public override void Dispose(bool disposeEventBinding)
		{
			if(this.Equals(GlobalHelperModules.GlobalModule.Instance))
				 GlobalHelperModules.GlobalModule.Instance = null;	
			base.Dispose(disposeEventBinding);
		}

		/// <summary>
		/// NetOffice method: dispose instance and all child instances
		/// </summary>
		[Category("NetOffice"), CoreOverridden]
		public override void Dispose()
		{
			if(this.Equals(GlobalHelperModules.GlobalModule.Instance))
				 GlobalHelperModules.GlobalModule.Instance = null;
			base.Dispose();
		}

        #endregion

        #region Static CoClass Methods

        /// <summary>
        /// Returns all running Outlook.Application instances from the environment/system
        /// </summary>
        /// <returns>Outlook.Application sequence</returns>
        public static IDisposableSequence<Application> GetActiveInstances()
        {
            return Running.ProxyService.GetActiveInstances<Application>("Outlook", "Application");
        }

        /// <summary>
        /// Returns a running Outlook.Application instance from the environment/system
        /// </summary>
        /// <param name="throwExceptionIfNotFound">throw exception if unable to find an instance</param>
        /// <returns>Outlook.Application instance or null</returns>
        public static Application GetActiveInstance(bool throwExceptionIfNotFound = false)
        {
            return Running.ProxyService.GetActiveInstance<Application>("Outlook", "Application", throwExceptionIfNotFound);
        }

        #endregion

        #region Events

        /// <summary>
        /// SupportByVersion Outlook, 9,10,11,12,14,15,16
        /// </summary>
        private event Application_ItemSendEventHandler _ItemSendEvent;

		/// <summary>
		/// SupportByVersion Outlook 9 10 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff865076.aspx </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		public event Application_ItemSendEventHandler ItemSendEvent
		{
			add
			{
				CreateEventBridge();
				_ItemSendEvent += value;
			}
			remove
			{
				_ItemSendEvent -= value;
			}
		}

		/// <summary>
		/// SupportByVersion Outlook, 9,10,11,12,14,15,16
		/// </summary>
		private event Application_NewMailEventHandler _NewMailEvent;

		/// <summary>
		/// SupportByVersion Outlook 9 10 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff869202.aspx </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		public event Application_NewMailEventHandler NewMailEvent
		{
			add
			{
				CreateEventBridge();
				_NewMailEvent += value;
			}
			remove
			{
				_NewMailEvent -= value;
			}
		}

		/// <summary>
		/// SupportByVersion Outlook, 9,10,11,12,14,15,16
		/// </summary>
		private event Application_ReminderEventHandler _ReminderEvent;

		/// <summary>
		/// SupportByVersion Outlook 9 10 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff870058.aspx </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		public event Application_ReminderEventHandler ReminderEvent
		{
			add
			{
				CreateEventBridge();
				_ReminderEvent += value;
			}
			remove
			{
				_ReminderEvent -= value;
			}
		}

		/// <summary>
		/// SupportByVersion Outlook, 9,10,11,12,14,15,16
		/// </summary>
		private event Application_OptionsPagesAddEventHandler _OptionsPagesAddEvent;

		/// <summary>
		/// SupportByVersion Outlook 9 10 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff868446.aspx </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		public event Application_OptionsPagesAddEventHandler OptionsPagesAddEvent
		{
			add
			{
				CreateEventBridge();
				_OptionsPagesAddEvent += value;
			}
			remove
			{
				_OptionsPagesAddEvent -= value;
			}
		}

		/// <summary>
		/// SupportByVersion Outlook, 9,10,11,12,14,15,16
		/// </summary>
		private event Application_StartupEventHandler _StartupEvent;

		/// <summary>
		/// SupportByVersion Outlook 9 10 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff869298.aspx </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		public event Application_StartupEventHandler StartupEvent
		{
			add
			{
				CreateEventBridge();
				_StartupEvent += value;
			}
			remove
			{
				_StartupEvent -= value;
			}
		}

		/// <summary>
		/// SupportByVersion Outlook, 9,10,11,12,14,15,16
		/// </summary>
		private event Application_QuitEventHandler _QuitEvent;

		/// <summary>
		/// SupportByVersion Outlook 9 10 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff869760.aspx </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		public event Application_QuitEventHandler QuitEvent
		{
			add
			{
				CreateEventBridge();
				_QuitEvent += value;
			}
			remove
			{
				_QuitEvent -= value;
			}
		}

		/// <summary>
		/// SupportByVersion Outlook, 10,11,12,14,15,16
		/// </summary>
		private event Application_AdvancedSearchCompleteEventHandler _AdvancedSearchCompleteEvent;

		/// <summary>
		/// SupportByVersion Outlook 10 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff864775.aspx </remarks>
		[SupportByVersion("Outlook", 10,11,12,14,15,16)]
		public event Application_AdvancedSearchCompleteEventHandler AdvancedSearchCompleteEvent
		{
			add
			{
				CreateEventBridge();
				_AdvancedSearchCompleteEvent += value;
			}
			remove
			{
				_AdvancedSearchCompleteEvent -= value;
			}
		}

		/// <summary>
		/// SupportByVersion Outlook, 10,11,12,14,15,16
		/// </summary>
		private event Application_AdvancedSearchStoppedEventHandler _AdvancedSearchStoppedEvent;

		/// <summary>
		/// SupportByVersion Outlook 10 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff868266.aspx </remarks>
		[SupportByVersion("Outlook", 10,11,12,14,15,16)]
		public event Application_AdvancedSearchStoppedEventHandler AdvancedSearchStoppedEvent
		{
			add
			{
				CreateEventBridge();
				_AdvancedSearchStoppedEvent += value;
			}
			remove
			{
				_AdvancedSearchStoppedEvent -= value;
			}
		}

		/// <summary>
		/// SupportByVersion Outlook, 10,11,12,14,15,16
		/// </summary>
		private event Application_MAPILogonCompleteEventHandler _MAPILogonCompleteEvent;

		/// <summary>
		/// SupportByVersion Outlook 10 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff869443.aspx </remarks>
		[SupportByVersion("Outlook", 10,11,12,14,15,16)]
		public event Application_MAPILogonCompleteEventHandler MAPILogonCompleteEvent
		{
			add
			{
				CreateEventBridge();
				_MAPILogonCompleteEvent += value;
			}
			remove
			{
				_MAPILogonCompleteEvent -= value;
			}
		}

		/// <summary>
		/// SupportByVersion Outlook, 11,12,14,15,16
		/// </summary>
		private event Application_NewMailExEventHandler _NewMailExEvent;

		/// <summary>
		/// SupportByVersion Outlook 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff863686.aspx </remarks>
		[SupportByVersion("Outlook", 11,12,14,15,16)]
		public event Application_NewMailExEventHandler NewMailExEvent
		{
			add
			{
				CreateEventBridge();
				_NewMailExEvent += value;
			}
			remove
			{
				_NewMailExEvent -= value;
			}
		}

		/// <summary>
		/// SupportByVersion Outlook, 12,14,15,16
		/// </summary>
		private event Application_AttachmentContextMenuDisplayEventHandler _AttachmentContextMenuDisplayEvent;

		/// <summary>
		/// SupportByVersion Outlook 12 14 15,16
		/// </summary>
		[SupportByVersion("Outlook", 12,14,15,16)]
		public event Application_AttachmentContextMenuDisplayEventHandler AttachmentContextMenuDisplayEvent
		{
			add
			{
				CreateEventBridge();
				_AttachmentContextMenuDisplayEvent += value;
			}
			remove
			{
				_AttachmentContextMenuDisplayEvent -= value;
			}
		}

		/// <summary>
		/// SupportByVersion Outlook, 12,14,15,16
		/// </summary>
		private event Application_FolderContextMenuDisplayEventHandler _FolderContextMenuDisplayEvent;

		/// <summary>
		/// SupportByVersion Outlook 12 14 15,16
		/// </summary>
		[SupportByVersion("Outlook", 12,14,15,16)]
		public event Application_FolderContextMenuDisplayEventHandler FolderContextMenuDisplayEvent
		{
			add
			{
				CreateEventBridge();
				_FolderContextMenuDisplayEvent += value;
			}
			remove
			{
				_FolderContextMenuDisplayEvent -= value;
			}
		}

		/// <summary>
		/// SupportByVersion Outlook, 12,14,15,16
		/// </summary>
		private event Application_StoreContextMenuDisplayEventHandler _StoreContextMenuDisplayEvent;

		/// <summary>
		/// SupportByVersion Outlook 12 14 15,16
		/// </summary>
		[SupportByVersion("Outlook", 12,14,15,16)]
		public event Application_StoreContextMenuDisplayEventHandler StoreContextMenuDisplayEvent
		{
			add
			{
				CreateEventBridge();
				_StoreContextMenuDisplayEvent += value;
			}
			remove
			{
				_StoreContextMenuDisplayEvent -= value;
			}
		}

		/// <summary>
		/// SupportByVersion Outlook, 12,14,15,16
		/// </summary>
		private event Application_ShortcutContextMenuDisplayEventHandler _ShortcutContextMenuDisplayEvent;

		/// <summary>
		/// SupportByVersion Outlook 12 14 15,16
		/// </summary>
		[SupportByVersion("Outlook", 12,14,15,16)]
		public event Application_ShortcutContextMenuDisplayEventHandler ShortcutContextMenuDisplayEvent
		{
			add
			{
				CreateEventBridge();
				_ShortcutContextMenuDisplayEvent += value;
			}
			remove
			{
				_ShortcutContextMenuDisplayEvent -= value;
			}
		}

		/// <summary>
		/// SupportByVersion Outlook, 12,14,15,16
		/// </summary>
		private event Application_ViewContextMenuDisplayEventHandler _ViewContextMenuDisplayEvent;

		/// <summary>
		/// SupportByVersion Outlook 12 14 15,16
		/// </summary>
		[SupportByVersion("Outlook", 12,14,15,16)]
		public event Application_ViewContextMenuDisplayEventHandler ViewContextMenuDisplayEvent
		{
			add
			{
				CreateEventBridge();
				_ViewContextMenuDisplayEvent += value;
			}
			remove
			{
				_ViewContextMenuDisplayEvent -= value;
			}
		}

		/// <summary>
		/// SupportByVersion Outlook, 12,14,15,16
		/// </summary>
		private event Application_ItemContextMenuDisplayEventHandler _ItemContextMenuDisplayEvent;

		/// <summary>
		/// SupportByVersion Outlook 12 14 15,16
		/// </summary>
		[SupportByVersion("Outlook", 12,14,15,16)]
		public event Application_ItemContextMenuDisplayEventHandler ItemContextMenuDisplayEvent
		{
			add
			{
				CreateEventBridge();
				_ItemContextMenuDisplayEvent += value;
			}
			remove
			{
				_ItemContextMenuDisplayEvent -= value;
			}
		}

		/// <summary>
		/// SupportByVersion Outlook, 12,14,15,16
		/// </summary>
		private event Application_ContextMenuCloseEventHandler _ContextMenuCloseEvent;

		/// <summary>
		/// SupportByVersion Outlook 12 14 15,16
		/// </summary>
		[SupportByVersion("Outlook", 12,14,15,16)]
		public event Application_ContextMenuCloseEventHandler ContextMenuCloseEvent
		{
			add
			{
				CreateEventBridge();
				_ContextMenuCloseEvent += value;
			}
			remove
			{
				_ContextMenuCloseEvent -= value;
			}
		}

		/// <summary>
		/// SupportByVersion Outlook, 12,14,15,16
		/// </summary>
		private event Application_ItemLoadEventHandler _ItemLoadEvent;

		/// <summary>
		/// SupportByVersion Outlook 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff868544.aspx </remarks>
		[SupportByVersion("Outlook", 12,14,15,16)]
		public event Application_ItemLoadEventHandler ItemLoadEvent
		{
			add
			{
				CreateEventBridge();
				_ItemLoadEvent += value;
			}
			remove
			{
				_ItemLoadEvent -= value;
			}
		}

		/// <summary>
		/// SupportByVersion Outlook, 12,14,15,16
		/// </summary>
		private event Application_BeforeFolderSharingDialogEventHandler _BeforeFolderSharingDialogEvent;

		/// <summary>
		/// SupportByVersion Outlook 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff869543.aspx </remarks>
		[SupportByVersion("Outlook", 12,14,15,16)]
		public event Application_BeforeFolderSharingDialogEventHandler BeforeFolderSharingDialogEvent
		{
			add
			{
				CreateEventBridge();
				_BeforeFolderSharingDialogEvent += value;
			}
			remove
			{
				_BeforeFolderSharingDialogEvent -= value;
			}
		}

		#endregion
       
	    #region IEventBinding Member
        
		/// <summary>
        /// Creates active sink helper
        /// </summary>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public void CreateEventBridge()
        {
			if(false == Factory.Settings.EnableEvents)
				return;
	
			if (null != _connectPoint)
				return;
	
            if (null == _activeSinkId)
				_activeSinkId = SinkHelper.GetConnectionPoint(this, ref _connectPoint, Events.ApplicationEvents_SinkHelper.Id, Events.ApplicationEvents_10_SinkHelper.Id, Events.ApplicationEvents_11_SinkHelper.Id);


			if(Events.ApplicationEvents_SinkHelper.Id.Equals(_activeSinkId, StringComparison.InvariantCultureIgnoreCase))
			{
				_applicationEvents_SinkHelper = new Events.ApplicationEvents_SinkHelper(this, _connectPoint);
				return;
			}

			if(Events.ApplicationEvents_10_SinkHelper.Id.Equals(_activeSinkId, StringComparison.InvariantCultureIgnoreCase))
			{
				_applicationEvents_10_SinkHelper = new Events.ApplicationEvents_10_SinkHelper(this, _connectPoint);
				return;
			}

			if(Events.ApplicationEvents_11_SinkHelper.Id.Equals(_activeSinkId, StringComparison.InvariantCultureIgnoreCase))
			{
				_applicationEvents_11_SinkHelper = new Events.ApplicationEvents_11_SinkHelper(this, _connectPoint);
				return;
			} 
        }

        /// <summary>
        /// The instance use currently an event listener 
        /// </summary>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        public bool EventBridgeInitialized
        {
            get 
            {
                return (null != _connectPoint);
            }
        }

        /// <summary>
        ///  The instance has currently one or more event recipients 
        /// </summary>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        public bool HasEventRecipients()       
        {
			if(null == _thisType)
				_thisType = this.GetType();
					
			foreach (NetRuntimeSystem.Reflection.EventInfo item in _thisType.GetEvents())
			{
				MulticastDelegate eventDelegate = (MulticastDelegate) _thisType.GetType().GetField(item.Name, 
																			NetRuntimeSystem.Reflection.BindingFlags.NonPublic |
																			NetRuntimeSystem.Reflection.BindingFlags.Instance).GetValue(this);
					
				if( (null != eventDelegate) && (eventDelegate.GetInvocationList().Length > 0) )
					return false;
			}
				
			return false;
        }
        
        /// <summary>
        /// Target methods from its actual event recipients
        /// </summary>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        public Delegate[] GetEventRecipients(string eventName)
        {
			if(null == _thisType)
				_thisType = this.GetType();
             
            MulticastDelegate eventDelegate = (MulticastDelegate)_thisType.GetField(
                                                "_" + eventName + "Event",
                                                NetRuntimeSystem.Reflection.BindingFlags.Instance |
                                                NetRuntimeSystem.Reflection.BindingFlags.NonPublic).GetValue(this);

            if (null != eventDelegate)
            {
                Delegate[] delegates = eventDelegate.GetInvocationList();
                return delegates;
            }
            else
                return new Delegate[0];
        }
       
        /// <summary>
        /// Returns the current count of event recipients
        /// </summary>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        public int GetCountOfEventRecipients(string eventName)
        {
			if(null == _thisType)
				_thisType = this.GetType();
             
            MulticastDelegate eventDelegate = (MulticastDelegate)_thisType.GetField(
                                                "_" + eventName + "Event",
                                                NetRuntimeSystem.Reflection.BindingFlags.Instance |
                                                NetRuntimeSystem.Reflection.BindingFlags.NonPublic).GetValue(this);

            if (null != eventDelegate)
            {
                Delegate[] delegates = eventDelegate.GetInvocationList();
                return delegates.Length;
            }
            else
                return 0;
           }
        
        /// <summary>
        /// Raise an instance event
        /// </summary>
        /// <param name="eventName">name of the event without 'Event' at the end</param>
        /// <param name="paramsArray">custom arguments for the event</param>
        /// <returns>count of called event recipients</returns>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        public int RaiseCustomEvent(string eventName, ref object[] paramsArray)
		{
			if(null == _thisType)
				_thisType = this.GetType();
             
            MulticastDelegate eventDelegate = (MulticastDelegate)_thisType.GetField(
                                                "_" + eventName + "Event",
                                                NetRuntimeSystem.Reflection.BindingFlags.Instance |
                                                NetRuntimeSystem.Reflection.BindingFlags.NonPublic).GetValue(this);

            if (null != eventDelegate)
            {
                Delegate[] delegates = eventDelegate.GetInvocationList();
                foreach (var item in delegates)
                {
                    try
                    {
                        item.Method.Invoke(item.Target, paramsArray);
                    }
                    catch (NetRuntimeSystem.Exception exception)
                    {
                        Factory.Console.WriteException(exception);
                    }
                }
                return delegates.Length;
            }
            else
                return 0;
		}

        /// <summary>
        /// Stop listening events for the instance
        /// </summary>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        public void DisposeEventBridge()
        {
			if( null != _applicationEvents_SinkHelper)
			{
				_applicationEvents_SinkHelper.Dispose();
				_applicationEvents_SinkHelper = null;
			}
			if( null != _applicationEvents_10_SinkHelper)
			{
				_applicationEvents_10_SinkHelper.Dispose();
				_applicationEvents_10_SinkHelper = null;
			}
			if( null != _applicationEvents_11_SinkHelper)
			{
				_applicationEvents_11_SinkHelper.Dispose();
				_applicationEvents_11_SinkHelper = null;
			}

			_connectPoint = null;
		}
        
        #endregion

		#pragma warning restore
	}
}