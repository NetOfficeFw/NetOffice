using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice;
using NetOffice.Misc;

namespace NetOffice.WordApi
{

	#region Delegates

	#pragma warning disable
	public delegate void Document_NewEventHandler();
	public delegate void Document_OpenEventHandler();
	public delegate void Document_CloseEventHandler();
	public delegate void Document_SyncEventHandler(NetOffice.OfficeApi.Enums.MsoSyncEventType SyncEventType);
	public delegate void Document_XMLAfterInsertEventHandler(NetOffice.WordApi.XMLNode NewXMLNode, bool InUndoRedo);
	public delegate void Document_XMLBeforeDeleteEventHandler(NetOffice.WordApi.Range DeletedRange, NetOffice.WordApi.XMLNode OldXMLNode, bool InUndoRedo);
	public delegate void Document_ContentControlAfterAddEventHandler(NetOffice.WordApi.ContentControl NewContentControl, bool InUndoRedo);
	public delegate void Document_ContentControlBeforeDeleteEventHandler(NetOffice.WordApi.ContentControl OldContentControl, bool InUndoRedo);
	public delegate void Document_ContentControlOnExitEventHandler(NetOffice.WordApi.ContentControl ContentControl, ref bool Cancel);
	public delegate void Document_ContentControlOnEnterEventHandler(NetOffice.WordApi.ContentControl ContentControl);
	public delegate void Document_ContentControlBeforeStoreUpdateEventHandler(NetOffice.WordApi.ContentControl ContentControl, ref string Content);
	public delegate void Document_ContentControlBeforeContentUpdateEventHandler(NetOffice.WordApi.ContentControl ContentControl, ref string Content);
	public delegate void Document_BuildingBlockInsertEventHandler(NetOffice.WordApi.Range Range, string Name, string Category, string BlockType, string Template);
	#pragma warning restore

	#endregion

	///<summary>
	/// CoClass Document 
	/// SupportByVersion Word, 9,10,11,12,14,15,16
	/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff822963.aspx
	///</summary>
	[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
	[EntityTypeAttribute(EntityType.IsCoClass)]
	public class Document : _Document,IEventBinding
	{
		#pragma warning disable
		#region Fields
		
		private NetRuntimeSystem.Runtime.InteropServices.ComTypes.IConnectionPoint _connectPoint;
		private string _activeSinkId;
		private NetRuntimeSystem.Type _thisType;
		DocumentEvents_SinkHelper _documentEvents_SinkHelper;
		DocumentEvents2_SinkHelper _documentEvents2_SinkHelper;
	
		#endregion

		#region Type Information

        /// <summary>
        /// Instance Type
        /// </summary>
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
                    _type = typeof(Document);
                return _type;
            }
        }
        
        #endregion
        		
		#region Construction

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
		
		///<summary>
        /// Creates a new instance of Document 
        ///</summary>		
		public Document():base("Word.Document")
		{
			
		}
		
		///<summary>
        /// Creates a new instance of Document
        ///</summary>
        ///<param name="progId">registered ProgID</param>
		public Document(string progId):base(progId)
		{
			
		}

		#endregion

		#region Static CoClass Methods

		/// <summary>
        /// Returns all running Word.Document objects from the environment/system
        /// </summary>
        /// <returns>an Word.Document array</returns>
		public static NetOffice.WordApi.Document[] GetActiveInstances()
		{		
			IDisposableEnumeration proxyList = NetOffice.ProxyService.GetActiveInstances("Word","Document");
			NetRuntimeSystem.Collections.Generic.List<NetOffice.WordApi.Document> resultList = new NetRuntimeSystem.Collections.Generic.List<NetOffice.WordApi.Document>();
			foreach(object proxy in proxyList)
				resultList.Add( new NetOffice.WordApi.Document(null, proxy) );
			return resultList.ToArray();
		}

		/// <summary>
        /// Returns a running Word.Document object from the environment/system.
        /// </summary>
        /// <returns>an Word.Document object or null</returns>
		public static NetOffice.WordApi.Document GetActiveInstance()
		{
			object proxy  = NetOffice.ProxyService.GetActiveInstance("Word","Document", false);
			if(null != proxy)
				return new NetOffice.WordApi.Document(null, proxy);
			else
				return null;
		}

		/// <summary>
        /// Returns a running Word.Document object from the environment/system. 
        /// </summary>
	    /// <param name="throwOnError">throw an exception if no object was found</param>
        /// <returns>an Word.Document object or null</returns>
		public static NetOffice.WordApi.Document GetActiveInstance(bool throwOnError)
		{
			object proxy  = NetOffice.ProxyService.GetActiveInstance("Word","Document", throwOnError);
			if(null != proxy)
				return new NetOffice.WordApi.Document(null, proxy);
			else
				return null;
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
				_activeSinkId = SinkHelper.GetConnectionPoint(this, ref _connectPoint, DocumentEvents_SinkHelper.Id,DocumentEvents2_SinkHelper.Id);


			if(DocumentEvents_SinkHelper.Id.Equals(_activeSinkId, StringComparison.InvariantCultureIgnoreCase))
			{
				_documentEvents_SinkHelper = new DocumentEvents_SinkHelper(this, _connectPoint);
				return;
			}

			if(DocumentEvents2_SinkHelper.Id.Equals(_activeSinkId, StringComparison.InvariantCultureIgnoreCase))
			{
				_documentEvents2_SinkHelper = new DocumentEvents2_SinkHelper(this, _connectPoint);
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