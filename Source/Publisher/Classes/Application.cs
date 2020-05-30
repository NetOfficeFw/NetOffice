﻿using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;
using NetOffice.CollectionsGeneric;

namespace NetOffice.PublisherApi
{
	#region Delegates

	#pragma warning disable
	public delegate void Application_WindowActivateEventHandler(NetOffice.PublisherApi.Window wn);
	public delegate void Application_WindowDeactivateEventHandler(NetOffice.PublisherApi.Window wn);
	public delegate void Application_WindowPageChangeEventHandler(NetOffice.PublisherApi.View vw);
	public delegate void Application_QuitEventHandler();
	public delegate void Application_NewDocumentEventHandler(NetOffice.PublisherApi._Document doc);
	public delegate void Application_DocumentOpenEventHandler(NetOffice.PublisherApi._Document doc);
	public delegate void Application_DocumentBeforeCloseEventHandler(NetOffice.PublisherApi._Document doc, ref bool cancel);
	public delegate void Application_MailMergeAfterMergeEventHandler(NetOffice.PublisherApi._Document doc);
	public delegate void Application_MailMergeAfterRecordMergeEventHandler(NetOffice.PublisherApi._Document doc);
	public delegate void Application_MailMergeBeforeMergeEventHandler(NetOffice.PublisherApi._Document doc, Int32 startRecord, Int32 endRecord, ref bool cancel);
	public delegate void Application_MailMergeBeforeRecordMergeEventHandler(NetOffice.PublisherApi._Document doc, ref bool cancel);
	public delegate void Application_MailMergeDataSourceLoadEventHandler(NetOffice.PublisherApi._Document doc);
	public delegate void Application_MailMergeWizardSendToCustomEventHandler(NetOffice.PublisherApi._Document doc);
	public delegate void Application_MailMergeWizardStateChangeEventHandler(NetOffice.PublisherApi._Document doc, Int32 fromState);
	public delegate void Application_MailMergeDataSourceValidateEventHandler(NetOffice.PublisherApi._Document doc, ref bool handled);
	public delegate void Application_MailMergeInsertBarcodeEventHandler(NetOffice.PublisherApi._Document doc, ref bool okToInsert);
	public delegate void Application_MailMergeRecipientListCloseEventHandler(NetOffice.PublisherApi._Document doc);
	public delegate void Application_MailMergeGenerateBarcodeEventHandler(NetOffice.PublisherApi._Document doc, ref string bstrString);
	public delegate void Application_MailMergeWizardFollowUpCustomEventHandler(NetOffice.PublisherApi._Document doc);
	public delegate void Application_BeforePrintEventHandler(NetOffice.PublisherApi._Document doc, ref bool cancel);
	public delegate void Application_AfterPrintEventHandler(NetOffice.PublisherApi._Document doc);
	public delegate void Application_ShowCatalogUIEventHandler();
	public delegate void Application_HideCatalogUIEventHandler();
	#pragma warning restore

	#endregion

	/// <summary>
	/// CoClass Application 
	/// SupportByVersion Publisher, 14,15,16
	/// </summary>
	[SupportByVersion("Publisher", 14,15,16)]
	[EntityType(EntityType.IsCoClass), ComProgId("Publisher.Application"), ModuleProvider(typeof(GlobalHelperModules.GlobalModule))]
	[EventSink(typeof(Events.ApplicationEvents_SinkHelper))]
    [ComEventInterface(typeof(Events.ApplicationEvents))]
    public class Application : _Application, ICloneable<Application>, IEventBinding
	{
		#pragma warning disable

		#region Fields
		
		private NetRuntimeSystem.Runtime.InteropServices.ComTypes.IConnectionPoint _connectPoint;
		private string _activeSinkId;
        private static Type _type;
        private Events.ApplicationEvents_SinkHelper _applicationEvents_SinkHelper;
	
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

		/// <param name="factory">current used factory core</param>
		/// <param name="parentObject">object there has created the proxy</param>
		/// <param name="proxyShare">proxy share instead if com proxy</param>
		public Application(Core factory, ICOMObject parentObject, COMProxyShare proxyShare) : base(factory, parentObject, proxyShare)
		{
			_callQuitInDispose = true;
		}

		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
		public Application(Core factory, ICOMObject parentObject, object comProxy) : base(factory, parentObject, comProxy)
		{			
			GlobalHelperModules.GlobalModule.Instance = this;
		}

        ///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
		public Application(ICOMObject parentObject, object comProxy) : base(parentObject, comProxy)
		{		
			GlobalHelperModules.GlobalModule.Instance = this;
		}

		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Application(Core factory, ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(factory, parentObject, comProxy, comProxyType)
		{
			
		}

		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Application(ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(parentObject, comProxy, comProxyType)
		{
			
		}
		
		///<param name="replacedObject">object to replaced. replacedObject are not usable after this action</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Application(ICOMObject replacedObject) : base(replacedObject)
		{
			
		}
		
		/// <summary>
        ///creates a new instance of Application 
        /// </summary>		
		public Application():base("Publisher.Application")
		{	
			GlobalHelperModules.GlobalModule.Instance = this;
		}
		
		/// <summary>
        ///creates a new instance of Application
        /// </summary>
        ///<param name="progId">registered ProgID</param>
		public Application(string progId):base(progId)
		{			
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
        /// Returns all running Publisher.Application instances from the environment/system
        /// </summary>
        /// <returns>Publisher.Application sequence</returns>
        public static IDisposableSequence<Application> GetActiveInstances()
        {
            return Running.ProxyService.GetActiveInstances<Application>("Publisher", "Application");
        }

        /// <summary>
        /// Returns a running Publisher.Application instance from the environment/system
        /// </summary>
        /// <param name="throwExceptionIfNotFound">throw exception if unable to find an instance</param>
        /// <returns>Publisher.Application instance or null</returns>
        public static Application GetActiveInstance(bool throwExceptionIfNotFound = false)
        {
            return Running.ProxyService.GetActiveInstance<Application>("Publisher", "Application", throwExceptionIfNotFound);
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
		public event Application_WindowActivateEventHandler WindowActivateEvent
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
		public event Application_WindowDeactivateEventHandler WindowDeactivateEvent
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
		public event Application_WindowPageChangeEventHandler WindowPageChangeEvent
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
		/// SupportByVersion Publisher, 14,15,16
		/// </summary>
		private event Application_NewDocumentEventHandler _NewDocumentEvent;

		/// <summary>
		/// SupportByVersion Publisher 14 15 16
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		public event Application_NewDocumentEventHandler NewDocumentEvent
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
		public event Application_DocumentOpenEventHandler DocumentOpenEvent
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
		public event Application_DocumentBeforeCloseEventHandler DocumentBeforeCloseEvent
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
		public event Application_MailMergeAfterMergeEventHandler MailMergeAfterMergeEvent
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
		public event Application_MailMergeAfterRecordMergeEventHandler MailMergeAfterRecordMergeEvent
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
		public event Application_MailMergeBeforeMergeEventHandler MailMergeBeforeMergeEvent
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
		public event Application_MailMergeBeforeRecordMergeEventHandler MailMergeBeforeRecordMergeEvent
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
		public event Application_MailMergeDataSourceLoadEventHandler MailMergeDataSourceLoadEvent
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
		public event Application_MailMergeWizardSendToCustomEventHandler MailMergeWizardSendToCustomEvent
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
		public event Application_MailMergeWizardStateChangeEventHandler MailMergeWizardStateChangeEvent
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
		public event Application_MailMergeDataSourceValidateEventHandler MailMergeDataSourceValidateEvent
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
		public event Application_MailMergeInsertBarcodeEventHandler MailMergeInsertBarcodeEvent
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
		public event Application_MailMergeRecipientListCloseEventHandler MailMergeRecipientListCloseEvent
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
		public event Application_MailMergeGenerateBarcodeEventHandler MailMergeGenerateBarcodeEvent
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
		public event Application_MailMergeWizardFollowUpCustomEventHandler MailMergeWizardFollowUpCustomEvent
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
		public event Application_BeforePrintEventHandler BeforePrintEvent
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
		public event Application_AfterPrintEventHandler AfterPrintEvent
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
		public event Application_ShowCatalogUIEventHandler ShowCatalogUIEvent
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
		public event Application_HideCatalogUIEventHandler HideCatalogUIEvent
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
       
	    #region IEventBinding
        
		/// <summary>
        /// creates active sink helper
        /// </summary>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public void CreateEventBridge()
        {
			if(false == Factory.Settings.EnableEvents)
				return;
	
			if (null != _connectPoint)
				return;
	
            if (null == _activeSinkId)
				_activeSinkId = SinkHelper.GetConnectionPoint(this, ref _connectPoint, Events.ApplicationEvents_SinkHelper.Id);


			if(Events.ApplicationEvents_SinkHelper.Id.Equals(_activeSinkId, StringComparison.InvariantCultureIgnoreCase))
			{
				_applicationEvents_SinkHelper = new Events.ApplicationEvents_SinkHelper(this, _connectPoint);
				return;
			} 
        }

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
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        public void DisposeEventBridge()
        {
			if( null != _applicationEvents_SinkHelper)
			{
				_applicationEvents_SinkHelper.Dispose();
				_applicationEvents_SinkHelper = null;
			}

			_connectPoint = null;
		}

        #endregion

        #region ICloneable<Application>

        /// <summary>
        /// Creates a new Application that is a copy of the current instance
        /// </summary>
        /// <returns>A new Application that is a copy of this instance</returns>
        /// <exception cref="NetOffice.Exceptions.CloneException">An unexpected error occurred. See inner exception(s) for details.</exception>
        public new virtual Application Clone()
        {
            return base.Clone() as Application;
        }

        #endregion

        #pragma warning restore
    }
}