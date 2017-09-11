using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.MSHTMLApi
{
	#region Delegates

	#pragma warning disable
	public delegate void HTMLWindow2_onloadEventHandler();
	public delegate void HTMLWindow2_onunloadEventHandler();
	public delegate void HTMLWindow2_onhelpEventHandler();
	public delegate void HTMLWindow2_onfocusEventHandler();
	public delegate void HTMLWindow2_onblurEventHandler();
	public delegate void HTMLWindow2_onerrorEventHandler(string description, string url, Int32 line);
	public delegate void HTMLWindow2_onresizeEventHandler();
	public delegate void HTMLWindow2_onscrollEventHandler();
	public delegate void HTMLWindow2_onbeforeunloadEventHandler();
	public delegate void HTMLWindow2_onbeforeprintEventHandler();
	public delegate void HTMLWindow2_onafterprintEventHandler();
	#pragma warning restore

	#endregion

	/// <summary>
	/// CoClass HTMLWindow2 
	/// SupportByVersion MSHTML, 4
	/// </summary>
	[SupportByVersion("MSHTML", 4)]
	[EntityType(EntityType.IsCoClass)]
	[EventSink(typeof(Events.HTMLWindowEvents_SinkHelper))]
    [ComEventInterface(typeof(Events.HTMLWindowEvents))]
    public class HTMLWindow2 : DispHTMLWindow2, IEventBinding
	{
		#pragma warning disable

		#region Fields
		
		private NetRuntimeSystem.Runtime.InteropServices.ComTypes.IConnectionPoint _connectPoint;
		private string _activeSinkId;
        private static Type _type;
        private Events.HTMLWindowEvents_SinkHelper _hTMLWindowEvents_SinkHelper;
	
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
                    _type = typeof(HTMLWindow2);
                return _type;
            }
        }
        
        #endregion
        		
		#region Construction

		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
		public HTMLWindow2(Core factory, ICOMObject parentObject, object comProxy) : base(factory, parentObject, comProxy)
		{
			
		}

        ///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
		public HTMLWindow2(ICOMObject parentObject, object comProxy) : base(parentObject, comProxy)
		{
			
		}

		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public HTMLWindow2(Core factory, ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(factory, parentObject, comProxy, comProxyType)
		{
			
		}

		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public HTMLWindow2(ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(parentObject, comProxy, comProxyType)
		{
			
		}
		
		///<param name="replacedObject">object to replaced. replacedObject are not usable after this action</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public HTMLWindow2(ICOMObject replacedObject) : base(replacedObject)
		{
			
		}
		
		/// <summary>
        /// Creates a new instance of HTMLWindow2 
        /// </summary>		
		public HTMLWindow2():base("MSHTML.HTMLWindow2")
		{
			
		}
		
		/// <summary>
        /// Creates a new instance of HTMLWindow2
        /// </summary>
        ///<param name="progId">registered ProgID</param>
		public HTMLWindow2(string progId):base(progId)
		{
			
		}

		#endregion

		#region Static CoClass Methods
		#endregion

		#region Events

		/// <summary>
		/// SupportByVersion MSHTML, 4
		/// </summary>
		private event HTMLWindow2_onloadEventHandler _onloadEvent;

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public event HTMLWindow2_onloadEventHandler onloadEvent
		{
			add
			{
				CreateEventBridge();
				_onloadEvent += value;
			}
			remove
			{
				_onloadEvent -= value;
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML, 4
		/// </summary>
		private event HTMLWindow2_onunloadEventHandler _onunloadEvent;

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public event HTMLWindow2_onunloadEventHandler onunloadEvent
		{
			add
			{
				CreateEventBridge();
				_onunloadEvent += value;
			}
			remove
			{
				_onunloadEvent -= value;
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML, 4
		/// </summary>
		private event HTMLWindow2_onhelpEventHandler _onhelpEvent;

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public event HTMLWindow2_onhelpEventHandler onhelpEvent
		{
			add
			{
				CreateEventBridge();
				_onhelpEvent += value;
			}
			remove
			{
				_onhelpEvent -= value;
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML, 4
		/// </summary>
		private event HTMLWindow2_onfocusEventHandler _onfocusEvent;

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public event HTMLWindow2_onfocusEventHandler onfocusEvent
		{
			add
			{
				CreateEventBridge();
				_onfocusEvent += value;
			}
			remove
			{
				_onfocusEvent -= value;
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML, 4
		/// </summary>
		private event HTMLWindow2_onblurEventHandler _onblurEvent;

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public event HTMLWindow2_onblurEventHandler onblurEvent
		{
			add
			{
				CreateEventBridge();
				_onblurEvent += value;
			}
			remove
			{
				_onblurEvent -= value;
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML, 4
		/// </summary>
		private event HTMLWindow2_onerrorEventHandler _onerrorEvent;

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public event HTMLWindow2_onerrorEventHandler onerrorEvent
		{
			add
			{
				CreateEventBridge();
				_onerrorEvent += value;
			}
			remove
			{
				_onerrorEvent -= value;
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML, 4
		/// </summary>
		private event HTMLWindow2_onresizeEventHandler _onresizeEvent;

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public event HTMLWindow2_onresizeEventHandler onresizeEvent
		{
			add
			{
				CreateEventBridge();
				_onresizeEvent += value;
			}
			remove
			{
				_onresizeEvent -= value;
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML, 4
		/// </summary>
		private event HTMLWindow2_onscrollEventHandler _onscrollEvent;

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public event HTMLWindow2_onscrollEventHandler onscrollEvent
		{
			add
			{
				CreateEventBridge();
				_onscrollEvent += value;
			}
			remove
			{
				_onscrollEvent -= value;
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML, 4
		/// </summary>
		private event HTMLWindow2_onbeforeunloadEventHandler _onbeforeunloadEvent;

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public event HTMLWindow2_onbeforeunloadEventHandler onbeforeunloadEvent
		{
			add
			{
				CreateEventBridge();
				_onbeforeunloadEvent += value;
			}
			remove
			{
				_onbeforeunloadEvent -= value;
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML, 4
		/// </summary>
		private event HTMLWindow2_onbeforeprintEventHandler _onbeforeprintEvent;

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public event HTMLWindow2_onbeforeprintEventHandler onbeforeprintEvent
		{
			add
			{
				CreateEventBridge();
				_onbeforeprintEvent += value;
			}
			remove
			{
				_onbeforeprintEvent -= value;
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML, 4
		/// </summary>
		private event HTMLWindow2_onafterprintEventHandler _onafterprintEvent;

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public event HTMLWindow2_onafterprintEventHandler onafterprintEvent
		{
			add
			{
				CreateEventBridge();
				_onafterprintEvent += value;
			}
			remove
			{
				_onafterprintEvent -= value;
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
				_activeSinkId = SinkHelper.GetConnectionPoint(this, ref _connectPoint, Events.HTMLWindowEvents_SinkHelper.Id);


			if(Events.HTMLWindowEvents_SinkHelper.Id.Equals(_activeSinkId, StringComparison.InvariantCultureIgnoreCase))
			{
				_hTMLWindowEvents_SinkHelper = new Events.HTMLWindowEvents_SinkHelper(this, _connectPoint);
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
			if( null != _hTMLWindowEvents_SinkHelper)
			{
				_hTMLWindowEvents_SinkHelper.Dispose();
				_hTMLWindowEvents_SinkHelper = null;
			}

			_connectPoint = null;
		}
        
        #endregion

		#pragma warning restore
	}
}

