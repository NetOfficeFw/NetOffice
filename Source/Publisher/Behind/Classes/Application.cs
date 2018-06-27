using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;
using NetOffice.CollectionsGeneric;

namespace NetOffice.PublisherApi.Behind
{
	/// <summary>
	/// CoClass Application
	/// SupportByVersion Publisher, 14,15,16
	/// </summary>
	[SupportByVersion("Publisher", 14,15,16)]
	[EntityType(EntityType.IsCoClass), ComProgId("Publisher.Application"), ModuleProvider(typeof(NetOffice.PublisherApi.ModulesLegacy.ApplicationModule))]
    [ComEventContract(typeof(NetOffice.PublisherApi.EventContracts.ApplicationEvents))]
    [HasInteropCompatibilityClass(typeof(ApplicationClass))]
    public class Application : _Application, NetOffice.PublisherApi.Application, IAutomaticQuit
    {
		#pragma warning disable

		#region Fields

		private NetRuntimeSystem.Runtime.InteropServices.ComTypes.IConnectionPoint _connectPoint;
		private string _activeSinkId;
        private static Type _type;
        private NetOffice.PublisherApi.Behind.EventContracts.ApplicationEvents_SinkHelper _applicationEvents_SinkHelper;

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
                    _contractType = typeof(NetOffice.PublisherApi.Application);
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
        /// Stub Ctor, not intended to use
        /// </summary>
        public Application() : base()
        {

        }

        /// <summary>
        /// Creates a new instance of Application
        /// <param name="enableProxyService">try to get a running application first before create a new application</param>
        /// </summary>
        public Application(Core factory = null, bool enableProxyService = false) : base()
        {
            object proxy = null;
            if (enableProxyService)
            {
                proxy = ProxyService.GetActiveInstance("Publisher", "Application", false);
                if (null != proxy)
                {
                    CreateFromProxy(proxy, true);
                    FromProxyService = true;
                }
            }

            if(null == proxy)
            {
                CreateFromProgId("Publisher.Application", true);
            }

            _callQuitInDispose = null == ParentObject;
            Factory = null != factory ? factory : Core.Default;
            OnCreate();
            ModulesLegacy.ApplicationModule.Instance = this;
        }

        #endregion

        #region Properties

        /// <summary>
        /// Instance is created from an already running application
        /// </summary>
        [EditorBrowsable(EditorBrowsableState.Advanced)]
        public virtual bool FromProxyService { get; private set; }

        #endregion

        #region Static CoClass Methods

        /// <summary>
        /// Returns all running Publisher.Application instances from the environment/system
        /// </summary>
        /// <returns>Publisher.Application sequence</returns>
        public static IDisposableSequence<Application> GetActiveInstances()
        {
            return ProxyService.GetActiveInstances<Application>("Publisher", "Application");
        }

        /// <summary>
        /// Returns all running Publisher.Application instances from the environment/system that passed a predicate filter
        /// </summary>
        /// <param name="predicate">filter predicate</param>
        /// <returns>Publisher.Application sequence</returns>
        public static IDisposableSequence<Application> GetActiveInstances(Func<Application, bool> predicate)
        {
            return ProxyService.GetActiveInstances<Application>("Publisher", "Application", predicate);
        }

        /// <summary>
        /// Returns the count of running Publisher.Application instances that passed a predicate filter
        /// </summary>
        /// <returns>count of running application</returns>
        public static int GetActiveInstancesCount()
        {
            var sequence = ProxyService.GetActiveInstances<Application>("Publisher", "Application");
            int result = sequence.Count;
            sequence.Dispose();
            return result;
        }

        /// <summary>
        /// Returns the count of running Publisher.Application instances that passed a predicate filter
        /// </summary>
        /// <param name="predicate">filter predicate</param>
        /// <returns>count of running application</returns>
        public static int GetActiveInstancesCount(Func<Application, bool> predicate)
        {
            var sequence = ProxyService.GetActiveInstances<Application>("Publisher", "Application", predicate);
            int result = sequence.Count;
            sequence.Dispose();
            return result;
        }

        /// <summary>
        /// Returns a running Publisher.Application instance from the environment/system
        /// </summary>
        /// <param name="throwExceptionIfNotFound">throw exception if unable to find an instance</param>
        /// <returns>Publisher.Application instance or null(Nothing in Visual Basic)</returns>
        public static Application GetActiveInstance(bool throwExceptionIfNotFound = false)
        {
            return ProxyService.GetActiveInstance<Application>("Publisher", "Application", throwExceptionIfNotFound);
        }

        /// <summary>
        /// Returns first running Publisher.Application instance from the environment/system that passed a predicate filter
        /// </summary>
        /// <param name="predicate">filter predicate</param>
        /// <param name="throwExceptionIfNotFound">throw exception if unable to find an instance</param>
        /// <returns>Publisher.Application instance or null(Nothing in Visual Basic)</returns>
        /// <exception cref="ArgumentOutOfRangeException">occurs if no instance match and throwExceptionIfNotFound is set</exception>
        public static Application GetActiveInstance(Func<Application, bool> predicate, bool throwExceptionIfNotFound = false)
        {
            return ProxyService.GetActiveInstance<Application>("Publisher", "Application", predicate, throwExceptionIfNotFound);
        }

        #endregion

        #region Events

        /// <summary>
        /// SupportByVersion Publisher, 14,15,16
        /// </summary>
        private event Application_WindowActivateEventHandler _WindowActivateEvent;

		/// <summary>
		/// SupportByVersion Publisher 14 15 16
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		public virtual event Application_WindowActivateEventHandler WindowActivateEvent
		{
			add
			{
				CreateEventBridge();
				_WindowActivateEvent += value;
			}
			remove
			{
				_WindowActivateEvent -= value;
			}
		}

		/// <summary>
		/// SupportByVersion Publisher, 14,15,16
		/// </summary>
		private event Application_WindowDeactivateEventHandler _WindowDeactivateEvent;

		/// <summary>
		/// SupportByVersion Publisher 14 15 16
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		public virtual event Application_WindowDeactivateEventHandler WindowDeactivateEvent
		{
			add
			{
				CreateEventBridge();
				_WindowDeactivateEvent += value;
			}
			remove
			{
				_WindowDeactivateEvent -= value;
			}
		}

		/// <summary>
		/// SupportByVersion Publisher, 14,15,16
		/// </summary>
		private event Application_WindowPageChangeEventHandler _WindowPageChangeEvent;

		/// <summary>
		/// SupportByVersion Publisher 14 15 16
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		public virtual event Application_WindowPageChangeEventHandler WindowPageChangeEvent
		{
			add
			{
				CreateEventBridge();
				_WindowPageChangeEvent += value;
			}
			remove
			{
				_WindowPageChangeEvent -= value;
			}
		}

		/// <summary>
		/// SupportByVersion Publisher, 14,15,16
		/// </summary>
		private event Application_QuitEventHandler _QuitEvent;

		/// <summary>
		/// SupportByVersion Publisher 14 15 16
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		public virtual event Application_QuitEventHandler QuitEvent
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
		/// SupportByVersion Publisher, 14,15,16
		/// </summary>
		private event Application_NewDocumentEventHandler _NewDocumentEvent;

		/// <summary>
		/// SupportByVersion Publisher 14 15 16
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		public virtual event Application_NewDocumentEventHandler NewDocumentEvent
		{
			add
			{
				CreateEventBridge();
				_NewDocumentEvent += value;
			}
			remove
			{
				_NewDocumentEvent -= value;
			}
		}

		/// <summary>
		/// SupportByVersion Publisher, 14,15,16
		/// </summary>
		private event Application_DocumentOpenEventHandler _DocumentOpenEvent;

		/// <summary>
		/// SupportByVersion Publisher 14 15 16
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		public virtual event Application_DocumentOpenEventHandler DocumentOpenEvent
		{
			add
			{
				CreateEventBridge();
				_DocumentOpenEvent += value;
			}
			remove
			{
				_DocumentOpenEvent -= value;
			}
		}

		/// <summary>
		/// SupportByVersion Publisher, 14,15,16
		/// </summary>
		private event Application_DocumentBeforeCloseEventHandler _DocumentBeforeCloseEvent;

		/// <summary>
		/// SupportByVersion Publisher 14 15 16
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		public virtual event Application_DocumentBeforeCloseEventHandler DocumentBeforeCloseEvent
		{
			add
			{
				CreateEventBridge();
				_DocumentBeforeCloseEvent += value;
			}
			remove
			{
				_DocumentBeforeCloseEvent -= value;
			}
		}

		/// <summary>
		/// SupportByVersion Publisher, 14,15,16
		/// </summary>
		private event Application_MailMergeAfterMergeEventHandler _MailMergeAfterMergeEvent;

		/// <summary>
		/// SupportByVersion Publisher 14 15 16
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		public virtual event Application_MailMergeAfterMergeEventHandler MailMergeAfterMergeEvent
		{
			add
			{
				CreateEventBridge();
				_MailMergeAfterMergeEvent += value;
			}
			remove
			{
				_MailMergeAfterMergeEvent -= value;
			}
		}

		/// <summary>
		/// SupportByVersion Publisher, 14,15,16
		/// </summary>
		private event Application_MailMergeAfterRecordMergeEventHandler _MailMergeAfterRecordMergeEvent;

		/// <summary>
		/// SupportByVersion Publisher 14 15 16
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		public virtual event Application_MailMergeAfterRecordMergeEventHandler MailMergeAfterRecordMergeEvent
		{
			add
			{
				CreateEventBridge();
				_MailMergeAfterRecordMergeEvent += value;
			}
			remove
			{
				_MailMergeAfterRecordMergeEvent -= value;
			}
		}

		/// <summary>
		/// SupportByVersion Publisher, 14,15,16
		/// </summary>
		private event Application_MailMergeBeforeMergeEventHandler _MailMergeBeforeMergeEvent;

		/// <summary>
		/// SupportByVersion Publisher 14 15 16
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		public virtual event Application_MailMergeBeforeMergeEventHandler MailMergeBeforeMergeEvent
		{
			add
			{
				CreateEventBridge();
				_MailMergeBeforeMergeEvent += value;
			}
			remove
			{
				_MailMergeBeforeMergeEvent -= value;
			}
		}

		/// <summary>
		/// SupportByVersion Publisher, 14,15,16
		/// </summary>
		private event Application_MailMergeBeforeRecordMergeEventHandler _MailMergeBeforeRecordMergeEvent;

		/// <summary>
		/// SupportByVersion Publisher 14 15 16
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		public virtual event Application_MailMergeBeforeRecordMergeEventHandler MailMergeBeforeRecordMergeEvent
		{
			add
			{
				CreateEventBridge();
				_MailMergeBeforeRecordMergeEvent += value;
			}
			remove
			{
				_MailMergeBeforeRecordMergeEvent -= value;
			}
		}

		/// <summary>
		/// SupportByVersion Publisher, 14,15,16
		/// </summary>
		private event Application_MailMergeDataSourceLoadEventHandler _MailMergeDataSourceLoadEvent;

		/// <summary>
		/// SupportByVersion Publisher 14 15 16
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		public virtual event Application_MailMergeDataSourceLoadEventHandler MailMergeDataSourceLoadEvent
		{
			add
			{
				CreateEventBridge();
				_MailMergeDataSourceLoadEvent += value;
			}
			remove
			{
				_MailMergeDataSourceLoadEvent -= value;
			}
		}

		/// <summary>
		/// SupportByVersion Publisher, 14,15,16
		/// </summary>
		private event Application_MailMergeWizardSendToCustomEventHandler _MailMergeWizardSendToCustomEvent;

		/// <summary>
		/// SupportByVersion Publisher 14 15 16
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		public virtual event Application_MailMergeWizardSendToCustomEventHandler MailMergeWizardSendToCustomEvent
		{
			add
			{
				CreateEventBridge();
				_MailMergeWizardSendToCustomEvent += value;
			}
			remove
			{
				_MailMergeWizardSendToCustomEvent -= value;
			}
		}

		/// <summary>
		/// SupportByVersion Publisher, 14,15,16
		/// </summary>
		private event Application_MailMergeWizardStateChangeEventHandler _MailMergeWizardStateChangeEvent;

		/// <summary>
		/// SupportByVersion Publisher 14 15 16
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		public virtual event Application_MailMergeWizardStateChangeEventHandler MailMergeWizardStateChangeEvent
		{
			add
			{
				CreateEventBridge();
				_MailMergeWizardStateChangeEvent += value;
			}
			remove
			{
				_MailMergeWizardStateChangeEvent -= value;
			}
		}

		/// <summary>
		/// SupportByVersion Publisher, 14,15,16
		/// </summary>
		private event Application_MailMergeDataSourceValidateEventHandler _MailMergeDataSourceValidateEvent;

		/// <summary>
		/// SupportByVersion Publisher 14 15 16
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		public virtual event Application_MailMergeDataSourceValidateEventHandler MailMergeDataSourceValidateEvent
		{
			add
			{
				CreateEventBridge();
				_MailMergeDataSourceValidateEvent += value;
			}
			remove
			{
				_MailMergeDataSourceValidateEvent -= value;
			}
		}

		/// <summary>
		/// SupportByVersion Publisher, 14,15,16
		/// </summary>
		private event Application_MailMergeInsertBarcodeEventHandler _MailMergeInsertBarcodeEvent;

		/// <summary>
		/// SupportByVersion Publisher 14 15 16
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		public virtual event Application_MailMergeInsertBarcodeEventHandler MailMergeInsertBarcodeEvent
		{
			add
			{
				CreateEventBridge();
				_MailMergeInsertBarcodeEvent += value;
			}
			remove
			{
				_MailMergeInsertBarcodeEvent -= value;
			}
		}

		/// <summary>
		/// SupportByVersion Publisher, 14,15,16
		/// </summary>
		private event Application_MailMergeRecipientListCloseEventHandler _MailMergeRecipientListCloseEvent;

		/// <summary>
		/// SupportByVersion Publisher 14 15 16
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		public virtual event Application_MailMergeRecipientListCloseEventHandler MailMergeRecipientListCloseEvent
		{
			add
			{
				CreateEventBridge();
				_MailMergeRecipientListCloseEvent += value;
			}
			remove
			{
				_MailMergeRecipientListCloseEvent -= value;
			}
		}

		/// <summary>
		/// SupportByVersion Publisher, 14,15,16
		/// </summary>
		private event Application_MailMergeGenerateBarcodeEventHandler _MailMergeGenerateBarcodeEvent;

		/// <summary>
		/// SupportByVersion Publisher 14 15 16
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		public virtual event Application_MailMergeGenerateBarcodeEventHandler MailMergeGenerateBarcodeEvent
		{
			add
			{
				CreateEventBridge();
				_MailMergeGenerateBarcodeEvent += value;
			}
			remove
			{
				_MailMergeGenerateBarcodeEvent -= value;
			}
		}

		/// <summary>
		/// SupportByVersion Publisher, 14,15,16
		/// </summary>
		private event Application_MailMergeWizardFollowUpCustomEventHandler _MailMergeWizardFollowUpCustomEvent;

		/// <summary>
		/// SupportByVersion Publisher 14 15 16
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		public virtual event Application_MailMergeWizardFollowUpCustomEventHandler MailMergeWizardFollowUpCustomEvent
		{
			add
			{
				CreateEventBridge();
				_MailMergeWizardFollowUpCustomEvent += value;
			}
			remove
			{
				_MailMergeWizardFollowUpCustomEvent -= value;
			}
		}

		/// <summary>
		/// SupportByVersion Publisher, 14,15,16
		/// </summary>
		private event Application_BeforePrintEventHandler _BeforePrintEvent;

		/// <summary>
		/// SupportByVersion Publisher 14 15 16
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		public virtual event Application_BeforePrintEventHandler BeforePrintEvent
		{
			add
			{
				CreateEventBridge();
				_BeforePrintEvent += value;
			}
			remove
			{
				_BeforePrintEvent -= value;
			}
		}

		/// <summary>
		/// SupportByVersion Publisher, 14,15,16
		/// </summary>
		private event Application_AfterPrintEventHandler _AfterPrintEvent;

		/// <summary>
		/// SupportByVersion Publisher 14 15 16
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		public virtual event Application_AfterPrintEventHandler AfterPrintEvent
		{
			add
			{
				CreateEventBridge();
				_AfterPrintEvent += value;
			}
			remove
			{
				_AfterPrintEvent -= value;
			}
		}

		/// <summary>
		/// SupportByVersion Publisher, 14,15,16
		/// </summary>
		private event Application_ShowCatalogUIEventHandler _ShowCatalogUIEvent;

		/// <summary>
		/// SupportByVersion Publisher 14 15 16
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		public virtual event Application_ShowCatalogUIEventHandler ShowCatalogUIEvent
		{
			add
			{
				CreateEventBridge();
				_ShowCatalogUIEvent += value;
			}
			remove
			{
				_ShowCatalogUIEvent -= value;
			}
		}

		/// <summary>
		/// SupportByVersion Publisher, 14,15,16
		/// </summary>
		private event Application_HideCatalogUIEventHandler _HideCatalogUIEvent;

		/// <summary>
		/// SupportByVersion Publisher 14 15 16
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		public virtual event Application_HideCatalogUIEventHandler HideCatalogUIEvent
		{
			add
			{
				CreateEventBridge();
				_HideCatalogUIEvent += value;
			}
			remove
			{
				_HideCatalogUIEvent -= value;
			}
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

        #region IEventBinding

        /// <summary>
        /// creates active sink helper
        /// </summary>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public virtual void CreateEventBridge()
        {
			if(false == Factory.Settings.EnableEvents)
				return;

			if (null != _connectPoint)
				return;

            if (null == _activeSinkId)
				_activeSinkId = SinkHelper.GetConnectionPoint(this, ref _connectPoint, NetOffice.PublisherApi.Behind.EventContracts.ApplicationEvents_SinkHelper.Id);


			if(NetOffice.PublisherApi.Behind.EventContracts.ApplicationEvents_SinkHelper.Id.Equals(_activeSinkId, StringComparison.InvariantCultureIgnoreCase))
			{
				_applicationEvents_SinkHelper = new NetOffice.PublisherApi.Behind.EventContracts.ApplicationEvents_SinkHelper(this, _connectPoint);
				return;
			}
        }

        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        public virtual bool EventBridgeInitialized
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
        public virtual bool HasEventRecipients()
        {
            return NetOffice.Events.CoClassEventReflector.HasEventRecipients(this, LateBindingApiWrapperType);
        }

        /// <summary>
        /// Instance has one or more event recipients
        /// </summary>
        /// <param name="eventName">name of the event</param>
        /// <returns></returns>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        public virtual bool HasEventRecipients(string eventName)
        {
            return NetOffice.Events.CoClassEventReflector.HasEventRecipients(this, LateBindingApiWrapperType, eventName);
        }

        /// <summary>
        /// Target methods from its actual event recipients
        /// </summary>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        public virtual Delegate[] GetEventRecipients(string eventName)
        {
            return NetOffice.Events.CoClassEventReflector.GetEventRecipients(this, LateBindingApiWrapperType, eventName);
        }

        /// <summary>
        /// Returns the current count of event recipients
        /// </summary>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        public virtual int GetCountOfEventRecipients(string eventName)
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
        public virtual int RaiseCustomEvent(string eventName, ref object[] paramsArray)
		{
            return NetOffice.Events.CoClassEventReflector.RaiseCustomEvent(this, LateBindingApiWrapperType, eventName, ref paramsArray);
		}
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        public virtual void DisposeEventBridge()
        {
			if( null != _applicationEvents_SinkHelper)
			{
				_applicationEvents_SinkHelper.Dispose();
				_applicationEvents_SinkHelper = null;
			}

			_connectPoint = null;
		}

        #endregion

        #region IDisposable

        /// <summary>
        /// NetOffice method: dispose instance and all child instances
        /// </summary>
        /// <param name="disposeEventBinding">dispose event exported proxies with one or more event recipients</param>
        [Category("NetOffice"), CoreOverridden]
        public override void Dispose(bool disposeEventBinding)
        {
            if (this.Equals(ModulesLegacy.ApplicationModule.Instance))
                ModulesLegacy.ApplicationModule.Instance = null;
            base.Dispose(disposeEventBinding);
        }

        /// <summary>
        /// NetOffice method: dispose instance and all child instances
        /// </summary>
        [Category("NetOffice"), CoreOverridden]
        public override void Dispose()
        {
            if (this.Equals(ModulesLegacy.ApplicationModule.Instance))
                ModulesLegacy.ApplicationModule.Instance = null;
            base.Dispose();
        }

        #endregion

        #region ICloneable<Application>

        /// <summary>
        /// Creates a new Application that is a copy of the current instance
        /// </summary>
        /// <returns>A new Application that is a copy of this instance</returns>
        /// <exception cref="CloneException">An unexpected error occured. See inner exception(s) for details.</exception>
        public new virtual NetOffice.PublisherApi.Application DeepCopy()
        {
            return base.Clone() as NetOffice.PublisherApi.Application;
        }

        #endregion

        #pragma warning restore
    }
}