using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;
using NetOffice.CollectionsGeneric;

namespace NetOffice.WordApi
{
	#region Delegates

	#pragma warning disable
	public delegate void Document_NewEventHandler();
	public delegate void Document_OpenEventHandler();
	public delegate void Document_CloseEventHandler();
	public delegate void Document_SyncEventHandler(NetOffice.OfficeApi.Enums.MsoSyncEventType syncEventType);
	public delegate void Document_XMLAfterInsertEventHandler(NetOffice.WordApi.XMLNode nNewXMLNode, bool inUndoRedo);
	public delegate void Document_XMLBeforeDeleteEventHandler(NetOffice.WordApi.Range deletedRange, NetOffice.WordApi.XMLNode oldXMLNode, bool inUndoRedo);
	public delegate void Document_ContentControlAfterAddEventHandler(NetOffice.WordApi.ContentControl newContentControl, bool inUndoRedo);
	public delegate void Document_ContentControlBeforeDeleteEventHandler(NetOffice.WordApi.ContentControl oldContentControl, bool inUndoRedo);
	public delegate void Document_ContentControlOnExitEventHandler(NetOffice.WordApi.ContentControl contentControl, ref bool cancel);
	public delegate void Document_ContentControlOnEnterEventHandler(NetOffice.WordApi.ContentControl contentControl);
	public delegate void Document_ContentControlBeforeStoreUpdateEventHandler(NetOffice.WordApi.ContentControl contentControl, ref string content);
	public delegate void Document_ContentControlBeforeContentUpdateEventHandler(NetOffice.WordApi.ContentControl contentControl, ref string content);
	public delegate void Document_BuildingBlockInsertEventHandler(NetOffice.WordApi.Range range, string name, string category, string blockType, string template);
	#pragma warning restore

	#endregion

	/// <summary>
	/// CoClass Document 
	/// SupportByVersion Word, 9,10,11,12,14,15,16
	/// </summary>
	/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff822963.aspx </remarks>
	[SupportByVersion("Word", 9,10,11,12,14,15,16)]
	[EntityType(EntityType.IsCoClass)]
	[EventSink(typeof(Events.DocumentEvents_SinkHelper), typeof(Events.DocumentEvents2_SinkHelper))]
    [ComEventInterface(typeof(Events.DocumentEvents), typeof(Events.DocumentEvents2))]
    public class Document : _Document, IEventBinding
	{
		#pragma warning disable

		#region Fields
		
		private NetRuntimeSystem.Runtime.InteropServices.ComTypes.IConnectionPoint _connectPoint;
		private string _activeSinkId;
        private static Type _type;
        private Events.DocumentEvents_SinkHelper _documentEvents_SinkHelper;
		private Events.DocumentEvents2_SinkHelper _documentEvents2_SinkHelper;
	
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
                    _type = typeof(Document);
                return _type;
            }
        }
        
        #endregion
        		
		#region Ctro

		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
		public Document(Core factory, ICOMObject parentObject, object comProxy) : base(factory, parentObject, comProxy)
		{
			
		}

        ///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
		public Document(ICOMObject parentObject, object comProxy) : base(parentObject, comProxy)
		{
			
		}

		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Document(Core factory, ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(factory, parentObject, comProxy, comProxyType)
		{
			
		}

		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Document(ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(parentObject, comProxy, comProxyType)
		{
			
		}
		
		///<param name="replacedObject">object to replaced. replacedObject are not usable after this action</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Document(ICOMObject replacedObject) : base(replacedObject)
		{
			
		}
		
		/// <summary>
        /// Creates a new instance of Document 
        /// </summary>		
		public Document():base("Word.Document")
		{
			
		}
		
		/// <summary>
        /// Creates a new instance of Document
        /// </summary>
        ///<param name="progId">registered ProgID</param>
		public Document(string progId):base(progId)
		{
			
		}

        #endregion

        #region Static CoClass Methods

        /// <summary>
        /// Returns all running Word.Document instances from the environment/system
        /// </summary>
        /// <returns>Word.Document sequence</returns>
        public static IDisposableSequence<Application> GetActiveInstances()
        {
            return Running.ProxyService.GetActiveInstances<Application>("Word", "Document");
        }

        /// <summary>
        /// Returns a running Word.Document instance from the environment/system
        /// </summary>
        /// <param name="throwExceptionIfNotFound">throw exception if unable to find an instance</param>
        /// <returns>Word.Document instance or null</returns>
        public static Application GetActiveInstance(bool throwExceptionIfNotFound = false)
        {
            return Running.ProxyService.GetActiveInstance<Application>("Word", "Document", throwExceptionIfNotFound);
        }

        #endregion

        #region Events

        /// <summary>
        /// SupportByVersion Word, 9,10,11,12,14,15,16
        /// </summary>
        private event Document_NewEventHandler _NewEvent;

		/// <summary>
		/// SupportByVersion Word 9 10 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff837882.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public event Document_NewEventHandler NewEvent
		{
			add
			{
				CreateEventBridge();
				_NewEvent += value;
			}
			remove
			{
				_NewEvent -= value;
			}
		}

		/// <summary>
		/// SupportByVersion Word, 9,10,11,12,14,15,16
		/// </summary>
		private event Document_OpenEventHandler _OpenEvent;

		/// <summary>
		/// SupportByVersion Word 9 10 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff821870.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public event Document_OpenEventHandler OpenEvent
		{
			add
			{
				CreateEventBridge();
				_OpenEvent += value;
			}
			remove
			{
				_OpenEvent -= value;
			}
		}

		/// <summary>
		/// SupportByVersion Word, 9,10,11,12,14,15,16
		/// </summary>
		private event Document_CloseEventHandler _CloseEvent;

		/// <summary>
		/// SupportByVersion Word 9 10 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff821142.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public event Document_CloseEventHandler CloseEvent
		{
			add
			{
				CreateEventBridge();
				_CloseEvent += value;
			}
			remove
			{
				_CloseEvent -= value;
			}
		}

		/// <summary>
		/// SupportByVersion Word, 11,12,14,15,16
		/// </summary>
		private event Document_SyncEventHandler _SyncEvent;

		/// <summary>
		/// SupportByVersion Word 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff838305.aspx </remarks>
		[SupportByVersion("Word", 11,12,14,15,16)]
		public event Document_SyncEventHandler SyncEvent
		{
			add
			{
				CreateEventBridge();
				_SyncEvent += value;
			}
			remove
			{
				_SyncEvent -= value;
			}
		}

		/// <summary>
		/// SupportByVersion Word, 11,12,14,15,16
		/// </summary>
		private event Document_XMLAfterInsertEventHandler _XMLAfterInsertEvent;

		/// <summary>
		/// SupportByVersion Word 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff197579.aspx </remarks>
		[SupportByVersion("Word", 11,12,14,15,16)]
		public event Document_XMLAfterInsertEventHandler XMLAfterInsertEvent
		{
			add
			{
				CreateEventBridge();
				_XMLAfterInsertEvent += value;
			}
			remove
			{
				_XMLAfterInsertEvent -= value;
			}
		}

		/// <summary>
		/// SupportByVersion Word, 11,12,14,15,16
		/// </summary>
		private event Document_XMLBeforeDeleteEventHandler _XMLBeforeDeleteEvent;

		/// <summary>
		/// SupportByVersion Word 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff191971.aspx </remarks>
		[SupportByVersion("Word", 11,12,14,15,16)]
		public event Document_XMLBeforeDeleteEventHandler XMLBeforeDeleteEvent
		{
			add
			{
				CreateEventBridge();
				_XMLBeforeDeleteEvent += value;
			}
			remove
			{
				_XMLBeforeDeleteEvent -= value;
			}
		}

		/// <summary>
		/// SupportByVersion Word, 12,14,15,16
		/// </summary>
		private event Document_ContentControlAfterAddEventHandler _ContentControlAfterAddEvent;

		/// <summary>
		/// SupportByVersion Word 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff834876.aspx </remarks>
		[SupportByVersion("Word", 12,14,15,16)]
		public event Document_ContentControlAfterAddEventHandler ContentControlAfterAddEvent
		{
			add
			{
				CreateEventBridge();
				_ContentControlAfterAddEvent += value;
			}
			remove
			{
				_ContentControlAfterAddEvent -= value;
			}
		}

		/// <summary>
		/// SupportByVersion Word, 12,14,15,16
		/// </summary>
		private event Document_ContentControlBeforeDeleteEventHandler _ContentControlBeforeDeleteEvent;

		/// <summary>
		/// SupportByVersion Word 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff835805.aspx </remarks>
		[SupportByVersion("Word", 12,14,15,16)]
		public event Document_ContentControlBeforeDeleteEventHandler ContentControlBeforeDeleteEvent
		{
			add
			{
				CreateEventBridge();
				_ContentControlBeforeDeleteEvent += value;
			}
			remove
			{
				_ContentControlBeforeDeleteEvent -= value;
			}
		}

		/// <summary>
		/// SupportByVersion Word, 12,14,15,16
		/// </summary>
		private event Document_ContentControlOnExitEventHandler _ContentControlOnExitEvent;

		/// <summary>
		/// SupportByVersion Word 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff191963.aspx </remarks>
		[SupportByVersion("Word", 12,14,15,16)]
		public event Document_ContentControlOnExitEventHandler ContentControlOnExitEvent
		{
			add
			{
				CreateEventBridge();
				_ContentControlOnExitEvent += value;
			}
			remove
			{
				_ContentControlOnExitEvent -= value;
			}
		}

		/// <summary>
		/// SupportByVersion Word, 12,14,15,16
		/// </summary>
		private event Document_ContentControlOnEnterEventHandler _ContentControlOnEnterEvent;

		/// <summary>
		/// SupportByVersion Word 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff196332.aspx </remarks>
		[SupportByVersion("Word", 12,14,15,16)]
		public event Document_ContentControlOnEnterEventHandler ContentControlOnEnterEvent
		{
			add
			{
				CreateEventBridge();
				_ContentControlOnEnterEvent += value;
			}
			remove
			{
				_ContentControlOnEnterEvent -= value;
			}
		}

		/// <summary>
		/// SupportByVersion Word, 12,14,15,16
		/// </summary>
		private event Document_ContentControlBeforeStoreUpdateEventHandler _ContentControlBeforeStoreUpdateEvent;

		/// <summary>
		/// SupportByVersion Word 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff835822.aspx </remarks>
		[SupportByVersion("Word", 12,14,15,16)]
		public event Document_ContentControlBeforeStoreUpdateEventHandler ContentControlBeforeStoreUpdateEvent
		{
			add
			{
				CreateEventBridge();
				_ContentControlBeforeStoreUpdateEvent += value;
			}
			remove
			{
				_ContentControlBeforeStoreUpdateEvent -= value;
			}
		}

		/// <summary>
		/// SupportByVersion Word, 12,14,15,16
		/// </summary>
		private event Document_ContentControlBeforeContentUpdateEventHandler _ContentControlBeforeContentUpdateEvent;

		/// <summary>
		/// SupportByVersion Word 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff192622.aspx </remarks>
		[SupportByVersion("Word", 12,14,15,16)]
		public event Document_ContentControlBeforeContentUpdateEventHandler ContentControlBeforeContentUpdateEvent
		{
			add
			{
				CreateEventBridge();
				_ContentControlBeforeContentUpdateEvent += value;
			}
			remove
			{
				_ContentControlBeforeContentUpdateEvent -= value;
			}
		}

		/// <summary>
		/// SupportByVersion Word, 12,14,15,16
		/// </summary>
		private event Document_BuildingBlockInsertEventHandler _BuildingBlockInsertEvent;

		/// <summary>
		/// SupportByVersion Word 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff197904.aspx </remarks>
		[SupportByVersion("Word", 12,14,15,16)]
		public event Document_BuildingBlockInsertEventHandler BuildingBlockInsertEvent
		{
			add
			{
				CreateEventBridge();
				_BuildingBlockInsertEvent += value;
			}
			remove
			{
				_BuildingBlockInsertEvent -= value;
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
			if(false == Factory.Settings.EnableEvents)
				return;
	
			if (null != _connectPoint)
				return;
	
            if (null == _activeSinkId)
				_activeSinkId = SinkHelper.GetConnectionPoint(this, ref _connectPoint, Events.DocumentEvents_SinkHelper.Id, Events.DocumentEvents2_SinkHelper.Id);


			if(Events.DocumentEvents_SinkHelper.Id.Equals(_activeSinkId, StringComparison.InvariantCultureIgnoreCase))
			{
				_documentEvents_SinkHelper = new Events.DocumentEvents_SinkHelper(this, _connectPoint);
				return;
			}

			if(Events.DocumentEvents2_SinkHelper.Id.Equals(_activeSinkId, StringComparison.InvariantCultureIgnoreCase))
			{
				_documentEvents2_SinkHelper = new Events.DocumentEvents2_SinkHelper(this, _connectPoint);
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
			if( null != _documentEvents_SinkHelper)
			{
				_documentEvents_SinkHelper.Dispose();
				_documentEvents_SinkHelper = null;
			}
			if( null != _documentEvents2_SinkHelper)
			{
				_documentEvents2_SinkHelper.Dispose();
				_documentEvents2_SinkHelper = null;
			}

			_connectPoint = null;
		}
        
        #endregion

		#pragma warning restore
	}
}

