using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.PublisherApi
{
	#region Delegates

	#pragma warning disable
	public delegate void Document_OpenEventHandler();
	public delegate void Document_BeforeCloseEventHandler(ref bool cancel);
	public delegate void Document_ShapesAddedEventHandler();
	public delegate void Document_WizardAfterChangeEventHandler();
	public delegate void Document_ShapesRemovedEventHandler();
	public delegate void Document_UndoEventHandler();
	public delegate void Document_RedoEventHandler();
	#pragma warning restore

	#endregion

	/// <summary>
	/// CoClass Document 
	/// SupportByVersion Publisher, 14,15,16
	/// </summary>
	[SupportByVersion("Publisher", 14,15,16)]
	[EntityType(EntityType.IsCoClass)]
	[EventSink(typeof(Events.DocumentEvents_SinkHelper))]
    [ComEventInterface(typeof(Events.DocumentEvents))]
    public class Document : _Document, IEventBinding
	{
		#pragma warning disable

		#region Fields
		
		private NetRuntimeSystem.Runtime.InteropServices.ComTypes.IConnectionPoint _connectPoint;
		private string _activeSinkId;
        private static Type _type;
        private Events.DocumentEvents_SinkHelper _documentEvents_SinkHelper;
	
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
		
		/// <summary>
        ///creates a new instance of Document 
        /// </summary>		
		public Document():base("Publisher.Document")
		{
			
		}
		
		/// <summary>
        ///creates a new instance of Document
        /// </summary>
        ///<param name="progId">registered ProgID</param>
		public Document(string progId):base(progId)
		{
			
		}

		#endregion

		#region Static CoClass Methods
		#endregion

		#region Events

		/// <summary>
		/// SupportByVersion Publisher, 14,15,16
		/// </summary>
		private event Document_OpenEventHandler _OpenEvent;

		/// <summary>
		/// SupportByVersion Publisher 14 15 16
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
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
		/// SupportByVersion Publisher, 14,15,16
		/// </summary>
		private event Document_BeforeCloseEventHandler _BeforeCloseEvent;

		/// <summary>
		/// SupportByVersion Publisher 14 15 16
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		public event Document_BeforeCloseEventHandler BeforeCloseEvent
		{
			add
			{
				CreateEventBridge();
				_BeforeCloseEvent += value;
			}
			remove
			{
				_BeforeCloseEvent -= value;
			}
		}

		/// <summary>
		/// SupportByVersion Publisher, 14,15,16
		/// </summary>
		private event Document_ShapesAddedEventHandler _ShapesAddedEvent;

		/// <summary>
		/// SupportByVersion Publisher 14 15 16
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		public event Document_ShapesAddedEventHandler ShapesAddedEvent
		{
			add
			{
				CreateEventBridge();
				_ShapesAddedEvent += value;
			}
			remove
			{
				_ShapesAddedEvent -= value;
			}
		}

		/// <summary>
		/// SupportByVersion Publisher, 14,15,16
		/// </summary>
		private event Document_WizardAfterChangeEventHandler _WizardAfterChangeEvent;

		/// <summary>
		/// SupportByVersion Publisher 14 15 16
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		public event Document_WizardAfterChangeEventHandler WizardAfterChangeEvent
		{
			add
			{
				CreateEventBridge();
				_WizardAfterChangeEvent += value;
			}
			remove
			{
				_WizardAfterChangeEvent -= value;
			}
		}

		/// <summary>
		/// SupportByVersion Publisher, 14,15,16
		/// </summary>
		private event Document_ShapesRemovedEventHandler _ShapesRemovedEvent;

		/// <summary>
		/// SupportByVersion Publisher 14 15 16
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		public event Document_ShapesRemovedEventHandler ShapesRemovedEvent
		{
			add
			{
				CreateEventBridge();
				_ShapesRemovedEvent += value;
			}
			remove
			{
				_ShapesRemovedEvent -= value;
			}
		}

		/// <summary>
		/// SupportByVersion Publisher, 14,15,16
		/// </summary>
		private event Document_UndoEventHandler _UndoEvent;

		/// <summary>
		/// SupportByVersion Publisher 14 15 16
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		public event Document_UndoEventHandler UndoEvent
		{
			add
			{
				CreateEventBridge();
				_UndoEvent += value;
			}
			remove
			{
				_UndoEvent -= value;
			}
		}

		/// <summary>
		/// SupportByVersion Publisher, 14,15,16
		/// </summary>
		private event Document_RedoEventHandler _RedoEvent;

		/// <summary>
		/// SupportByVersion Publisher 14 15 16
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		public event Document_RedoEventHandler RedoEvent
		{
			add
			{
				CreateEventBridge();
				_RedoEvent += value;
			}
			remove
			{
				_RedoEvent -= value;
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
				_activeSinkId = SinkHelper.GetConnectionPoint(this, ref _connectPoint, Events.DocumentEvents_SinkHelper.Id);


			if(Events.DocumentEvents_SinkHelper.Id.Equals(_activeSinkId, StringComparison.InvariantCultureIgnoreCase))
			{
				_documentEvents_SinkHelper = new Events.DocumentEvents_SinkHelper(this, _connectPoint);
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
			if( null != _documentEvents_SinkHelper)
			{
				_documentEvents_SinkHelper.Dispose();
				_documentEvents_SinkHelper = null;
			}

			_connectPoint = null;
		}
        
        #endregion

		#pragma warning restore
	}
}

