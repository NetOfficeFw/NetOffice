using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.CoreServices;
using NetOffice.Attributes;
using NetOffice.CollectionsGeneric;

namespace NetOffice.OutlookApi.Behind
{
	/// <summary>
	/// CoClass Application
	/// SupportByVersion Outlook, 9,10,11,12,14,15,16
	/// </summary>
	/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff866895.aspx </remarks>
	[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
	[EntityType(EntityType.IsCoClass), ComProgId("Outlook.Application"), ModuleProvider(typeof(ModulesLegacy.ApplicationModule))]
    [ComEventContract(typeof(NetOffice.OutlookApi.EventContracts.ApplicationEvents), typeof(NetOffice.OutlookApi.EventContracts.ApplicationEvents_10), typeof(NetOffice.OutlookApi.EventContracts.ApplicationEvents_11))]
    [HasInteropCompatibilityClass(typeof(ApplicationClass))]
    public class Application : _Application, NetOffice.OutlookApi.Application, IApplicationVersionProvider, IAutomaticQuit
    {
		#pragma warning disable

		#region Fields

		private NetRuntimeSystem.Runtime.InteropServices.ComTypes.IConnectionPoint _connectPoint;
		private string _activeSinkId;
        private static Type _type;
        private EventContracts.ApplicationEvents_SinkHelper _applicationEvents_SinkHelper;
        private EventContracts.ApplicationEvents_10_SinkHelper _applicationEvents_10_SinkHelper;
        private EventContracts.ApplicationEvents_11_SinkHelper _applicationEvents_11_SinkHelper;

        private bool _versionRequested;
        private object _cachedVersion;
        private object _chachedVersionLock = new object();

        #endregion

        #region Type Information

        /// <summary>
        /// Contract Type
        /// </summary>
        [EditorBrowsable(EditorBrowsableState.Advanced), Browsable(false), Category("NetOffice"), CoreOverridden]
        public override Type ContractType
        {
            get
            {
                if(null == _contractType)
                    _contractType = typeof(NetOffice.OutlookApi.Application);
                return _contractType;
            }
        }
        private static Type _contractType;


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

        /// <summary>
        /// Type Cache
        /// </summary>
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

		#region Ctor

		/// <summary>
		/// Stub Ctor, not indented to use
		/// </summary>
		public Application() : base()
		{

		}

        /// <summary>
        /// Creates a new instance of the class
        /// <param name="factory">factory core instead of default core</param>
        /// <param name="tryProxyServiceFirst">try to get a running application first before create a new application</param>
        /// </summary>
        public Application(Core factory = null, bool tryProxyServiceFirst = false) : base()
        {
            object proxy = null;
            if (tryProxyServiceFirst)
            {
                proxy = ProxyService.GetActiveInstance("Outlook", "Application", false);
                if (null != proxy)
                {
                    CreateFromProxy(proxy, true);
                    FromProxyService = true;
                }
            }

            if(null == proxy)
            {
                CreateFromProgId("Outlook.Application", true);
            }

            Factory = null != factory ? factory : Core.Default;
            TryRequestVersion();
            RegisterAsApplicationVersionProvider();
            OnCreate();
            _isInitialized = true;
        }

        #endregion

        #region Properties

        /// <summary>
        /// Instance is created from an already running application
        /// </summary>
        [EditorBrowsable(EditorBrowsableState.Advanced)]
        public bool FromProxyService { get; private set; }

        #endregion

        #region ICloneable<Application>

        /// <summary>
        /// Creates a new Application that is a copy of the current instance
        /// </summary>
        /// <returns>A new Application that is a copy of this instance</returns>
        /// <exception cref="CloneException">An unexpected error occured. See inner exception(s) for details.</exception>
        public new virtual NetOffice.OutlookApi.Application DeepCopy()
        {
            return base.Clone() as NetOffice.OutlookApi.Application;
        }

        #endregion

        #region IAutomaticQuit

        /// <summary>
        /// Determines Quit method want be called while disposing if NetOffice.Settings.EnableAutomaticQuit is true.
        /// Default is true when instance has no parent object and its not a cloned instance, otherwise false.
        /// </summary>
        bool IAutomaticQuit.Enabled
        {

            get
            {
                return _callQuitInDispose;
            }
            set
            {
                _callQuitInDispose = value;
            }
        }

        #endregion

        #region IApplicationVersionProvider

        string IApplicationVersionProvider.Name
        {
            get
            {
                return "Microsoft Outlook";
            }
        }

        string IApplicationVersionProvider.ComponentName
        {
            get
            {
                return "NetOffice.OutlookApi";
            }
        }

        /// <summary>
        /// Request version information on demand and cache to call the remote server only 1x times
        /// </summary>
        object IApplicationVersionProvider.Version
        {
            get
            {
                lock (_chachedVersionLock)
                {
                    if (null == _cachedVersion)
                    {
                        _cachedVersion = TryVersionPropertyGet();
                    }
                }
                return _cachedVersion;
            }
        }

        bool IApplicationVersionProvider.VersionRequested
        {
            get
            {
                return _versionRequested;
            }
        }

        void IApplicationVersionProvider.TryRequestVersion()
        {
            _cachedVersion = TryVersionPropertyGet();
        }

        /// <summary>
        /// Try get version information without fail
        /// </summary>
        /// <returns></returns>
        private object TryVersionPropertyGet()
        {
            try
            {
                if (null != _proxyShare)
                    return Invoker.PropertyGet(this, "Version");
                else
                    return null;
            }
            catch
            {
                return null;
            }
            finally
            {
                if (null != _proxyShare)
                    _versionRequested = true;
            }
        }

        #endregion

        #region IEventBinding

        /// <summary>
        /// Creates active sink helper
        /// </summary>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        public void CreateEventBridge()
        {
            if (false == Factory.Settings.EnableEvents)
                return;

            if (null != _connectPoint)
                return;

            if (null == _activeSinkId)
                _activeSinkId = SinkHelper.GetConnectionPoint(this, ref _connectPoint, NetOffice.OutlookApi.Behind.EventContracts.ApplicationEvents_SinkHelper.Id, NetOffice.OutlookApi.Behind.EventContracts.ApplicationEvents_10_SinkHelper.Id, NetOffice.OutlookApi.Behind.EventContracts.ApplicationEvents_11_SinkHelper.Id);


            if (NetOffice.OutlookApi.Behind.EventContracts.ApplicationEvents_SinkHelper.Id.Equals(_activeSinkId, StringComparison.InvariantCultureIgnoreCase))
            {
                _applicationEvents_SinkHelper = new NetOffice.OutlookApi.Behind.EventContracts.ApplicationEvents_SinkHelper(this, _connectPoint);
                return;
            }

            if (NetOffice.OutlookApi.Behind.EventContracts.ApplicationEvents_10_SinkHelper.Id.Equals(_activeSinkId, StringComparison.InvariantCultureIgnoreCase))
            {
                _applicationEvents_10_SinkHelper = new NetOffice.OutlookApi.Behind.EventContracts.ApplicationEvents_10_SinkHelper(this, _connectPoint);
                return;
            }

            if (NetOffice.OutlookApi.Behind.EventContracts.ApplicationEvents_11_SinkHelper.Id.Equals(_activeSinkId, StringComparison.InvariantCultureIgnoreCase))
            {
                _applicationEvents_11_SinkHelper = new NetOffice.OutlookApi.Behind.EventContracts.ApplicationEvents_11_SinkHelper(this, _connectPoint);
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
        /// Instance has one or more event recipients
        /// </summary>
        /// <returns>true if one or more event is active, otherwise false</returns>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        public bool HasEventRecipients()
        {
            return NetOffice.Events.CoClassEventReflector.HasEventRecipients(this, LateBindingApiWrapperType);
        }

        /// <summary>
        /// Instance has one or more event recipients
        /// </summary>
        /// <param name="eventName">name of the event</param>
        /// <returns></returns>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        public bool HasEventRecipients(string eventName)
        {
            return NetOffice.Events.CoClassEventReflector.HasEventRecipients(this, LateBindingApiWrapperType, eventName);
        }

        /// <summary>
        /// Target methods from its actual event recipients
        /// </summary>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        public Delegate[] GetEventRecipients(string eventName)
        {
            return NetOffice.Events.CoClassEventReflector.GetEventRecipients(this, LateBindingApiWrapperType, eventName);
        }

        /// <summary>
        /// Returns the current count of event recipients
        /// </summary>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        public int GetCountOfEventRecipients(string eventName)
        {
            return NetOffice.Events.CoClassEventReflector.GetCountOfEventRecipients(this, LateBindingApiWrapperType, eventName);
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
            return NetOffice.Events.CoClassEventReflector.RaiseCustomEvent(this, LateBindingApiWrapperType, eventName, ref paramsArray);
        }
        /// <summary>
        /// Stop listening events for the instance
        /// </summary>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        public void DisposeEventBridge()
        {
            if (null != _applicationEvents_SinkHelper)
            {
                _applicationEvents_SinkHelper.Dispose();
                _applicationEvents_SinkHelper = null;
            }
            if (null != _applicationEvents_10_SinkHelper)
            {
                _applicationEvents_10_SinkHelper.Dispose();
                _applicationEvents_10_SinkHelper = null;
            }
            if (null != _applicationEvents_11_SinkHelper)
            {
                _applicationEvents_11_SinkHelper.Dispose();
                _applicationEvents_11_SinkHelper = null;
            }

            _connectPoint = null;
        }

        #endregion

        #region Static CoClass Methods

        /// <summary>
        /// Returns all running Outlook.Application instances from the environment/system
        /// </summary>
        /// <returns>Outlook.Application sequence</returns>
        public static IDisposableSequence<Application> GetActiveInstances()
        {
            return ProxyService.GetActiveInstances<Application>("Outlook", "Application");
        }

        /// <summary>
        /// Returns all running Outlook.Application instances from the environment/system that passed a predicate filter
        /// </summary>
        /// <param name="predicate">filter predicate</param>
        /// <returns>Outlook.Application sequence</returns>
        public static IDisposableSequence<Application> GetActiveInstances(Func<Application, bool> predicate)
        {
            return ProxyService.GetActiveInstances<Application>("Outlook", "Application", predicate);
        }

        /// <summary>
        /// Returns the count of running Outlook.Application instances that passed a predicate filter
        /// </summary>
        /// <returns>count of running application</returns>
        public static int GetActiveInstancesCount()
        {
            var sequence = ProxyService.GetActiveInstances<Application>("Outlook", "Application");
            int result = sequence.Count;
            sequence.Dispose();
            return result;
        }

        /// <summary>
        /// Returns the count of running Outlook.Application instances that passed a predicate filter
        /// </summary>
        /// <param name="predicate">filter predicate</param>
        /// <returns>count of running application</returns>
        public static int GetActiveInstancesCount(Func<Application, bool> predicate)
        {
            var sequence = ProxyService.GetActiveInstances<Application>("Outlook", "Application", predicate);
            int result = sequence.Count;
            sequence.Dispose();
            return result;
        }

        /// <summary>
        /// Returns a running Outlook.Application instance from the environment/system
        /// </summary>
        /// <param name="throwExceptionIfNotFound">throw exception if unable to find an instance</param>
        /// <returns>Outlook.Application instance or null(Nothing in Visual Basic)</returns>
        /// <exception cref="ArgumentOutOfRangeException">occurs if no instance match and throwExceptionIfNotFound is set</exception>
        public static Application GetActiveInstance(bool throwExceptionIfNotFound = false)
        {
            return ProxyService.GetActiveInstance<Application>("Outlook", "Application", throwExceptionIfNotFound);
        }

        /// <summary>
        /// Returns first running Outlook.Application instance from the environment/system that passed a predicate filter
        /// </summary>
        /// <param name="predicate">filter predicate</param>
        /// <param name="throwExceptionIfNotFound">throw exception if unable to find an instance</param>
        /// <returns>Outlook.Application instance or null(Nothing in Visual Basic)</returns>
        /// <exception cref="ArgumentOutOfRangeException">occurs if no instance match and throwExceptionIfNotFound is set</exception>
        public static Application GetActiveInstance(Func<Application, bool> predicate, bool throwExceptionIfNotFound = false)
        {
            return ProxyService.GetActiveInstance<Application>("Outlook", "Application", predicate, throwExceptionIfNotFound);
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

        #pragma warning restore
    }
}
