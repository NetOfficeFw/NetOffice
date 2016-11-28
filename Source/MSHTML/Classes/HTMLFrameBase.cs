using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice;
using NetOffice.Misc;

namespace NetOffice.MSHTMLApi
{

	#region Delegates

	#pragma warning disable
	public delegate void HTMLFrameBase_onhelpEventHandler();
	public delegate void HTMLFrameBase_onclickEventHandler();
	public delegate void HTMLFrameBase_ondblclickEventHandler();
	public delegate void HTMLFrameBase_onkeypressEventHandler();
	public delegate void HTMLFrameBase_onkeydownEventHandler();
	public delegate void HTMLFrameBase_onkeyupEventHandler();
	public delegate void HTMLFrameBase_onmouseoutEventHandler();
	public delegate void HTMLFrameBase_onmouseoverEventHandler();
	public delegate void HTMLFrameBase_onmousemoveEventHandler();
	public delegate void HTMLFrameBase_onmousedownEventHandler();
	public delegate void HTMLFrameBase_onmouseupEventHandler();
	public delegate void HTMLFrameBase_onselectstartEventHandler();
	public delegate void HTMLFrameBase_onfilterchangeEventHandler();
	public delegate void HTMLFrameBase_ondragstartEventHandler();
	public delegate void HTMLFrameBase_onbeforeupdateEventHandler();
	public delegate void HTMLFrameBase_onafterupdateEventHandler();
	public delegate void HTMLFrameBase_onerrorupdateEventHandler();
	public delegate void HTMLFrameBase_onrowexitEventHandler();
	public delegate void HTMLFrameBase_onrowenterEventHandler();
	public delegate void HTMLFrameBase_ondatasetchangedEventHandler();
	public delegate void HTMLFrameBase_ondataavailableEventHandler();
	public delegate void HTMLFrameBase_ondatasetcompleteEventHandler();
	public delegate void HTMLFrameBase_onlosecaptureEventHandler();
	public delegate void HTMLFrameBase_onpropertychangeEventHandler();
	public delegate void HTMLFrameBase_onscrollEventHandler();
	public delegate void HTMLFrameBase_onfocusEventHandler();
	public delegate void HTMLFrameBase_onblurEventHandler();
	public delegate void HTMLFrameBase_onresizeEventHandler();
	public delegate void HTMLFrameBase_ondragEventHandler();
	public delegate void HTMLFrameBase_ondragendEventHandler();
	public delegate void HTMLFrameBase_ondragenterEventHandler();
	public delegate void HTMLFrameBase_ondragoverEventHandler();
	public delegate void HTMLFrameBase_ondragleaveEventHandler();
	public delegate void HTMLFrameBase_ondropEventHandler();
	public delegate void HTMLFrameBase_onbeforecutEventHandler();
	public delegate void HTMLFrameBase_oncutEventHandler();
	public delegate void HTMLFrameBase_onbeforecopyEventHandler();
	public delegate void HTMLFrameBase_oncopyEventHandler();
	public delegate void HTMLFrameBase_onbeforepasteEventHandler();
	public delegate void HTMLFrameBase_onpasteEventHandler();
	public delegate void HTMLFrameBase_oncontextmenuEventHandler();
	public delegate void HTMLFrameBase_onrowsdeleteEventHandler();
	public delegate void HTMLFrameBase_onrowsinsertedEventHandler();
	public delegate void HTMLFrameBase_oncellchangeEventHandler();
	public delegate void HTMLFrameBase_onreadystatechangeEventHandler();
	public delegate void HTMLFrameBase_onbeforeeditfocusEventHandler();
	public delegate void HTMLFrameBase_onlayoutcompleteEventHandler();
	public delegate void HTMLFrameBase_onpageEventHandler();
	public delegate void HTMLFrameBase_onbeforedeactivateEventHandler();
	public delegate void HTMLFrameBase_onbeforeactivateEventHandler();
	public delegate void HTMLFrameBase_onmoveEventHandler();
	public delegate void HTMLFrameBase_oncontrolselectEventHandler();
	public delegate void HTMLFrameBase_onmovestartEventHandler();
	public delegate void HTMLFrameBase_onmoveendEventHandler();
	public delegate void HTMLFrameBase_onresizestartEventHandler();
	public delegate void HTMLFrameBase_onresizeendEventHandler();
	public delegate void HTMLFrameBase_onmouseenterEventHandler();
	public delegate void HTMLFrameBase_onmouseleaveEventHandler();
	public delegate void HTMLFrameBase_onmousewheelEventHandler();
	public delegate void HTMLFrameBase_onactivateEventHandler();
	public delegate void HTMLFrameBase_ondeactivateEventHandler();
	public delegate void HTMLFrameBase_onfocusinEventHandler();
	public delegate void HTMLFrameBase_onfocusoutEventHandler();
	#pragma warning restore

	#endregion

	///<summary>
	/// CoClass HTMLFrameBase 
	/// SupportByVersion MSHTML, 4
	///</summary>
	[SupportByVersionAttribute("MSHTML", 4)]
	[EntityTypeAttribute(EntityType.IsCoClass)]
	public class HTMLFrameBase : DispHTMLFrameBase,IEventBinding
	{
		#pragma warning disable
		#region Fields
		
		private NetRuntimeSystem.Runtime.InteropServices.ComTypes.IConnectionPoint _connectPoint;
		private string _activeSinkId;
		private NetRuntimeSystem.Type _thisType;
		HTMLControlElementEvents_SinkHelper _hTMLControlElementEvents_SinkHelper;
	
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
                    _type = typeof(HTMLFrameBase);
                return _type;
            }
        }
        
        #endregion
        		
		#region Construction

		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
		public HTMLFrameBase(Core factory, ICOMObject parentObject, object comProxy) : base(factory, parentObject, comProxy)
		{
			
		}

        ///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
		public HTMLFrameBase(ICOMObject parentObject, object comProxy) : base(parentObject, comProxy)
		{
			
		}

		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public HTMLFrameBase(Core factory, ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(factory, parentObject, comProxy, comProxyType)
		{
			
		}

		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public HTMLFrameBase(ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(parentObject, comProxy, comProxyType)
		{
			
		}
		
		///<param name="replacedObject">object to replaced. replacedObject are not usable after this action</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public HTMLFrameBase(ICOMObject replacedObject) : base(replacedObject)
		{
			
		}
		
		///<summary>
        /// Creates a new instance of HTMLFrameBase 
        ///</summary>		
		public HTMLFrameBase():base("MSHTML.HTMLFrameBase")
		{
			
		}
		
		///<summary>
        /// Creates a new instance of HTMLFrameBase
        ///</summary>
        ///<param name="progId">registered ProgID</param>
		public HTMLFrameBase(string progId):base(progId)
		{
			
		}

		#endregion

		#region Static CoClass Methods

		/// <summary>
        /// Returns all running MSHTML.HTMLFrameBase objects from the environment/system
        /// </summary>
        /// <returns>an MSHTML.HTMLFrameBase array</returns>
		public static NetOffice.MSHTMLApi.HTMLFrameBase[] GetActiveInstances()
		{		
			IDisposableEnumeration proxyList = NetOffice.ProxyService.GetActiveInstances("MSHTML","HTMLFrameBase");
			NetRuntimeSystem.Collections.Generic.List<NetOffice.MSHTMLApi.HTMLFrameBase> resultList = new NetRuntimeSystem.Collections.Generic.List<NetOffice.MSHTMLApi.HTMLFrameBase>();
			foreach(object proxy in proxyList)
				resultList.Add( new NetOffice.MSHTMLApi.HTMLFrameBase(null, proxy) );
			return resultList.ToArray();
		}

		/// <summary>
        /// Returns a running MSHTML.HTMLFrameBase object from the environment/system.
        /// </summary>
        /// <returns>an MSHTML.HTMLFrameBase object or null</returns>
		public static NetOffice.MSHTMLApi.HTMLFrameBase GetActiveInstance()
		{
			object proxy  = NetOffice.ProxyService.GetActiveInstance("MSHTML","HTMLFrameBase", false);
			if(null != proxy)
				return new NetOffice.MSHTMLApi.HTMLFrameBase(null, proxy);
			else
				return null;
		}

		/// <summary>
        /// Returns a running MSHTML.HTMLFrameBase object from the environment/system. 
        /// </summary>
	    /// <param name="throwOnError">throw an exception if no object was found</param>
        /// <returns>an MSHTML.HTMLFrameBase object or null</returns>
		public static NetOffice.MSHTMLApi.HTMLFrameBase GetActiveInstance(bool throwOnError)
		{
			object proxy  = NetOffice.ProxyService.GetActiveInstance("MSHTML","HTMLFrameBase", throwOnError);
			if(null != proxy)
				return new NetOffice.MSHTMLApi.HTMLFrameBase(null, proxy);
			else
				return null;
		}
		#endregion

		#region Events

		/// <summary>
		/// SupportByVersion MSHTML, 4
		/// </summary>
		private event HTMLFrameBase_onhelpEventHandler _onhelpEvent;

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public event HTMLFrameBase_onhelpEventHandler onhelpEvent
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
		private event HTMLFrameBase_onclickEventHandler _onclickEvent;

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public event HTMLFrameBase_onclickEventHandler onclickEvent
		{
			add
			{
				CreateEventBridge();
				_onclickEvent += value;
			}
			remove
			{
				_onclickEvent -= value;
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML, 4
		/// </summary>
		private event HTMLFrameBase_ondblclickEventHandler _ondblclickEvent;

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public event HTMLFrameBase_ondblclickEventHandler ondblclickEvent
		{
			add
			{
				CreateEventBridge();
				_ondblclickEvent += value;
			}
			remove
			{
				_ondblclickEvent -= value;
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML, 4
		/// </summary>
		private event HTMLFrameBase_onkeypressEventHandler _onkeypressEvent;

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public event HTMLFrameBase_onkeypressEventHandler onkeypressEvent
		{
			add
			{
				CreateEventBridge();
				_onkeypressEvent += value;
			}
			remove
			{
				_onkeypressEvent -= value;
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML, 4
		/// </summary>
		private event HTMLFrameBase_onkeydownEventHandler _onkeydownEvent;

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public event HTMLFrameBase_onkeydownEventHandler onkeydownEvent
		{
			add
			{
				CreateEventBridge();
				_onkeydownEvent += value;
			}
			remove
			{
				_onkeydownEvent -= value;
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML, 4
		/// </summary>
		private event HTMLFrameBase_onkeyupEventHandler _onkeyupEvent;

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public event HTMLFrameBase_onkeyupEventHandler onkeyupEvent
		{
			add
			{
				CreateEventBridge();
				_onkeyupEvent += value;
			}
			remove
			{
				_onkeyupEvent -= value;
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML, 4
		/// </summary>
		private event HTMLFrameBase_onmouseoutEventHandler _onmouseoutEvent;

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public event HTMLFrameBase_onmouseoutEventHandler onmouseoutEvent
		{
			add
			{
				CreateEventBridge();
				_onmouseoutEvent += value;
			}
			remove
			{
				_onmouseoutEvent -= value;
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML, 4
		/// </summary>
		private event HTMLFrameBase_onmouseoverEventHandler _onmouseoverEvent;

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public event HTMLFrameBase_onmouseoverEventHandler onmouseoverEvent
		{
			add
			{
				CreateEventBridge();
				_onmouseoverEvent += value;
			}
			remove
			{
				_onmouseoverEvent -= value;
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML, 4
		/// </summary>
		private event HTMLFrameBase_onmousemoveEventHandler _onmousemoveEvent;

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public event HTMLFrameBase_onmousemoveEventHandler onmousemoveEvent
		{
			add
			{
				CreateEventBridge();
				_onmousemoveEvent += value;
			}
			remove
			{
				_onmousemoveEvent -= value;
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML, 4
		/// </summary>
		private event HTMLFrameBase_onmousedownEventHandler _onmousedownEvent;

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public event HTMLFrameBase_onmousedownEventHandler onmousedownEvent
		{
			add
			{
				CreateEventBridge();
				_onmousedownEvent += value;
			}
			remove
			{
				_onmousedownEvent -= value;
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML, 4
		/// </summary>
		private event HTMLFrameBase_onmouseupEventHandler _onmouseupEvent;

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public event HTMLFrameBase_onmouseupEventHandler onmouseupEvent
		{
			add
			{
				CreateEventBridge();
				_onmouseupEvent += value;
			}
			remove
			{
				_onmouseupEvent -= value;
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML, 4
		/// </summary>
		private event HTMLFrameBase_onselectstartEventHandler _onselectstartEvent;

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public event HTMLFrameBase_onselectstartEventHandler onselectstartEvent
		{
			add
			{
				CreateEventBridge();
				_onselectstartEvent += value;
			}
			remove
			{
				_onselectstartEvent -= value;
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML, 4
		/// </summary>
		private event HTMLFrameBase_onfilterchangeEventHandler _onfilterchangeEvent;

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public event HTMLFrameBase_onfilterchangeEventHandler onfilterchangeEvent
		{
			add
			{
				CreateEventBridge();
				_onfilterchangeEvent += value;
			}
			remove
			{
				_onfilterchangeEvent -= value;
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML, 4
		/// </summary>
		private event HTMLFrameBase_ondragstartEventHandler _ondragstartEvent;

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public event HTMLFrameBase_ondragstartEventHandler ondragstartEvent
		{
			add
			{
				CreateEventBridge();
				_ondragstartEvent += value;
			}
			remove
			{
				_ondragstartEvent -= value;
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML, 4
		/// </summary>
		private event HTMLFrameBase_onbeforeupdateEventHandler _onbeforeupdateEvent;

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public event HTMLFrameBase_onbeforeupdateEventHandler onbeforeupdateEvent
		{
			add
			{
				CreateEventBridge();
				_onbeforeupdateEvent += value;
			}
			remove
			{
				_onbeforeupdateEvent -= value;
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML, 4
		/// </summary>
		private event HTMLFrameBase_onafterupdateEventHandler _onafterupdateEvent;

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public event HTMLFrameBase_onafterupdateEventHandler onafterupdateEvent
		{
			add
			{
				CreateEventBridge();
				_onafterupdateEvent += value;
			}
			remove
			{
				_onafterupdateEvent -= value;
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML, 4
		/// </summary>
		private event HTMLFrameBase_onerrorupdateEventHandler _onerrorupdateEvent;

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public event HTMLFrameBase_onerrorupdateEventHandler onerrorupdateEvent
		{
			add
			{
				CreateEventBridge();
				_onerrorupdateEvent += value;
			}
			remove
			{
				_onerrorupdateEvent -= value;
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML, 4
		/// </summary>
		private event HTMLFrameBase_onrowexitEventHandler _onrowexitEvent;

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public event HTMLFrameBase_onrowexitEventHandler onrowexitEvent
		{
			add
			{
				CreateEventBridge();
				_onrowexitEvent += value;
			}
			remove
			{
				_onrowexitEvent -= value;
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML, 4
		/// </summary>
		private event HTMLFrameBase_onrowenterEventHandler _onrowenterEvent;

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public event HTMLFrameBase_onrowenterEventHandler onrowenterEvent
		{
			add
			{
				CreateEventBridge();
				_onrowenterEvent += value;
			}
			remove
			{
				_onrowenterEvent -= value;
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML, 4
		/// </summary>
		private event HTMLFrameBase_ondatasetchangedEventHandler _ondatasetchangedEvent;

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public event HTMLFrameBase_ondatasetchangedEventHandler ondatasetchangedEvent
		{
			add
			{
				CreateEventBridge();
				_ondatasetchangedEvent += value;
			}
			remove
			{
				_ondatasetchangedEvent -= value;
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML, 4
		/// </summary>
		private event HTMLFrameBase_ondataavailableEventHandler _ondataavailableEvent;

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public event HTMLFrameBase_ondataavailableEventHandler ondataavailableEvent
		{
			add
			{
				CreateEventBridge();
				_ondataavailableEvent += value;
			}
			remove
			{
				_ondataavailableEvent -= value;
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML, 4
		/// </summary>
		private event HTMLFrameBase_ondatasetcompleteEventHandler _ondatasetcompleteEvent;

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public event HTMLFrameBase_ondatasetcompleteEventHandler ondatasetcompleteEvent
		{
			add
			{
				CreateEventBridge();
				_ondatasetcompleteEvent += value;
			}
			remove
			{
				_ondatasetcompleteEvent -= value;
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML, 4
		/// </summary>
		private event HTMLFrameBase_onlosecaptureEventHandler _onlosecaptureEvent;

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public event HTMLFrameBase_onlosecaptureEventHandler onlosecaptureEvent
		{
			add
			{
				CreateEventBridge();
				_onlosecaptureEvent += value;
			}
			remove
			{
				_onlosecaptureEvent -= value;
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML, 4
		/// </summary>
		private event HTMLFrameBase_onpropertychangeEventHandler _onpropertychangeEvent;

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public event HTMLFrameBase_onpropertychangeEventHandler onpropertychangeEvent
		{
			add
			{
				CreateEventBridge();
				_onpropertychangeEvent += value;
			}
			remove
			{
				_onpropertychangeEvent -= value;
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML, 4
		/// </summary>
		private event HTMLFrameBase_onscrollEventHandler _onscrollEvent;

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public event HTMLFrameBase_onscrollEventHandler onscrollEvent
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
		private event HTMLFrameBase_onfocusEventHandler _onfocusEvent;

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public event HTMLFrameBase_onfocusEventHandler onfocusEvent
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
		private event HTMLFrameBase_onblurEventHandler _onblurEvent;

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public event HTMLFrameBase_onblurEventHandler onblurEvent
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
		private event HTMLFrameBase_onresizeEventHandler _onresizeEvent;

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public event HTMLFrameBase_onresizeEventHandler onresizeEvent
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
		private event HTMLFrameBase_ondragEventHandler _ondragEvent;

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public event HTMLFrameBase_ondragEventHandler ondragEvent
		{
			add
			{
				CreateEventBridge();
				_ondragEvent += value;
			}
			remove
			{
				_ondragEvent -= value;
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML, 4
		/// </summary>
		private event HTMLFrameBase_ondragendEventHandler _ondragendEvent;

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public event HTMLFrameBase_ondragendEventHandler ondragendEvent
		{
			add
			{
				CreateEventBridge();
				_ondragendEvent += value;
			}
			remove
			{
				_ondragendEvent -= value;
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML, 4
		/// </summary>
		private event HTMLFrameBase_ondragenterEventHandler _ondragenterEvent;

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public event HTMLFrameBase_ondragenterEventHandler ondragenterEvent
		{
			add
			{
				CreateEventBridge();
				_ondragenterEvent += value;
			}
			remove
			{
				_ondragenterEvent -= value;
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML, 4
		/// </summary>
		private event HTMLFrameBase_ondragoverEventHandler _ondragoverEvent;

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public event HTMLFrameBase_ondragoverEventHandler ondragoverEvent
		{
			add
			{
				CreateEventBridge();
				_ondragoverEvent += value;
			}
			remove
			{
				_ondragoverEvent -= value;
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML, 4
		/// </summary>
		private event HTMLFrameBase_ondragleaveEventHandler _ondragleaveEvent;

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public event HTMLFrameBase_ondragleaveEventHandler ondragleaveEvent
		{
			add
			{
				CreateEventBridge();
				_ondragleaveEvent += value;
			}
			remove
			{
				_ondragleaveEvent -= value;
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML, 4
		/// </summary>
		private event HTMLFrameBase_ondropEventHandler _ondropEvent;

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public event HTMLFrameBase_ondropEventHandler ondropEvent
		{
			add
			{
				CreateEventBridge();
				_ondropEvent += value;
			}
			remove
			{
				_ondropEvent -= value;
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML, 4
		/// </summary>
		private event HTMLFrameBase_onbeforecutEventHandler _onbeforecutEvent;

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public event HTMLFrameBase_onbeforecutEventHandler onbeforecutEvent
		{
			add
			{
				CreateEventBridge();
				_onbeforecutEvent += value;
			}
			remove
			{
				_onbeforecutEvent -= value;
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML, 4
		/// </summary>
		private event HTMLFrameBase_oncutEventHandler _oncutEvent;

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public event HTMLFrameBase_oncutEventHandler oncutEvent
		{
			add
			{
				CreateEventBridge();
				_oncutEvent += value;
			}
			remove
			{
				_oncutEvent -= value;
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML, 4
		/// </summary>
		private event HTMLFrameBase_onbeforecopyEventHandler _onbeforecopyEvent;

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public event HTMLFrameBase_onbeforecopyEventHandler onbeforecopyEvent
		{
			add
			{
				CreateEventBridge();
				_onbeforecopyEvent += value;
			}
			remove
			{
				_onbeforecopyEvent -= value;
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML, 4
		/// </summary>
		private event HTMLFrameBase_oncopyEventHandler _oncopyEvent;

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public event HTMLFrameBase_oncopyEventHandler oncopyEvent
		{
			add
			{
				CreateEventBridge();
				_oncopyEvent += value;
			}
			remove
			{
				_oncopyEvent -= value;
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML, 4
		/// </summary>
		private event HTMLFrameBase_onbeforepasteEventHandler _onbeforepasteEvent;

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public event HTMLFrameBase_onbeforepasteEventHandler onbeforepasteEvent
		{
			add
			{
				CreateEventBridge();
				_onbeforepasteEvent += value;
			}
			remove
			{
				_onbeforepasteEvent -= value;
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML, 4
		/// </summary>
		private event HTMLFrameBase_onpasteEventHandler _onpasteEvent;

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public event HTMLFrameBase_onpasteEventHandler onpasteEvent
		{
			add
			{
				CreateEventBridge();
				_onpasteEvent += value;
			}
			remove
			{
				_onpasteEvent -= value;
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML, 4
		/// </summary>
		private event HTMLFrameBase_oncontextmenuEventHandler _oncontextmenuEvent;

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public event HTMLFrameBase_oncontextmenuEventHandler oncontextmenuEvent
		{
			add
			{
				CreateEventBridge();
				_oncontextmenuEvent += value;
			}
			remove
			{
				_oncontextmenuEvent -= value;
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML, 4
		/// </summary>
		private event HTMLFrameBase_onrowsdeleteEventHandler _onrowsdeleteEvent;

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public event HTMLFrameBase_onrowsdeleteEventHandler onrowsdeleteEvent
		{
			add
			{
				CreateEventBridge();
				_onrowsdeleteEvent += value;
			}
			remove
			{
				_onrowsdeleteEvent -= value;
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML, 4
		/// </summary>
		private event HTMLFrameBase_onrowsinsertedEventHandler _onrowsinsertedEvent;

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public event HTMLFrameBase_onrowsinsertedEventHandler onrowsinsertedEvent
		{
			add
			{
				CreateEventBridge();
				_onrowsinsertedEvent += value;
			}
			remove
			{
				_onrowsinsertedEvent -= value;
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML, 4
		/// </summary>
		private event HTMLFrameBase_oncellchangeEventHandler _oncellchangeEvent;

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public event HTMLFrameBase_oncellchangeEventHandler oncellchangeEvent
		{
			add
			{
				CreateEventBridge();
				_oncellchangeEvent += value;
			}
			remove
			{
				_oncellchangeEvent -= value;
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML, 4
		/// </summary>
		private event HTMLFrameBase_onreadystatechangeEventHandler _onreadystatechangeEvent;

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public event HTMLFrameBase_onreadystatechangeEventHandler onreadystatechangeEvent
		{
			add
			{
				CreateEventBridge();
				_onreadystatechangeEvent += value;
			}
			remove
			{
				_onreadystatechangeEvent -= value;
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML, 4
		/// </summary>
		private event HTMLFrameBase_onbeforeeditfocusEventHandler _onbeforeeditfocusEvent;

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public event HTMLFrameBase_onbeforeeditfocusEventHandler onbeforeeditfocusEvent
		{
			add
			{
				CreateEventBridge();
				_onbeforeeditfocusEvent += value;
			}
			remove
			{
				_onbeforeeditfocusEvent -= value;
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML, 4
		/// </summary>
		private event HTMLFrameBase_onlayoutcompleteEventHandler _onlayoutcompleteEvent;

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public event HTMLFrameBase_onlayoutcompleteEventHandler onlayoutcompleteEvent
		{
			add
			{
				CreateEventBridge();
				_onlayoutcompleteEvent += value;
			}
			remove
			{
				_onlayoutcompleteEvent -= value;
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML, 4
		/// </summary>
		private event HTMLFrameBase_onpageEventHandler _onpageEvent;

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public event HTMLFrameBase_onpageEventHandler onpageEvent
		{
			add
			{
				CreateEventBridge();
				_onpageEvent += value;
			}
			remove
			{
				_onpageEvent -= value;
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML, 4
		/// </summary>
		private event HTMLFrameBase_onbeforedeactivateEventHandler _onbeforedeactivateEvent;

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public event HTMLFrameBase_onbeforedeactivateEventHandler onbeforedeactivateEvent
		{
			add
			{
				CreateEventBridge();
				_onbeforedeactivateEvent += value;
			}
			remove
			{
				_onbeforedeactivateEvent -= value;
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML, 4
		/// </summary>
		private event HTMLFrameBase_onbeforeactivateEventHandler _onbeforeactivateEvent;

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public event HTMLFrameBase_onbeforeactivateEventHandler onbeforeactivateEvent
		{
			add
			{
				CreateEventBridge();
				_onbeforeactivateEvent += value;
			}
			remove
			{
				_onbeforeactivateEvent -= value;
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML, 4
		/// </summary>
		private event HTMLFrameBase_onmoveEventHandler _onmoveEvent;

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public event HTMLFrameBase_onmoveEventHandler onmoveEvent
		{
			add
			{
				CreateEventBridge();
				_onmoveEvent += value;
			}
			remove
			{
				_onmoveEvent -= value;
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML, 4
		/// </summary>
		private event HTMLFrameBase_oncontrolselectEventHandler _oncontrolselectEvent;

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public event HTMLFrameBase_oncontrolselectEventHandler oncontrolselectEvent
		{
			add
			{
				CreateEventBridge();
				_oncontrolselectEvent += value;
			}
			remove
			{
				_oncontrolselectEvent -= value;
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML, 4
		/// </summary>
		private event HTMLFrameBase_onmovestartEventHandler _onmovestartEvent;

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public event HTMLFrameBase_onmovestartEventHandler onmovestartEvent
		{
			add
			{
				CreateEventBridge();
				_onmovestartEvent += value;
			}
			remove
			{
				_onmovestartEvent -= value;
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML, 4
		/// </summary>
		private event HTMLFrameBase_onmoveendEventHandler _onmoveendEvent;

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public event HTMLFrameBase_onmoveendEventHandler onmoveendEvent
		{
			add
			{
				CreateEventBridge();
				_onmoveendEvent += value;
			}
			remove
			{
				_onmoveendEvent -= value;
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML, 4
		/// </summary>
		private event HTMLFrameBase_onresizestartEventHandler _onresizestartEvent;

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public event HTMLFrameBase_onresizestartEventHandler onresizestartEvent
		{
			add
			{
				CreateEventBridge();
				_onresizestartEvent += value;
			}
			remove
			{
				_onresizestartEvent -= value;
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML, 4
		/// </summary>
		private event HTMLFrameBase_onresizeendEventHandler _onresizeendEvent;

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public event HTMLFrameBase_onresizeendEventHandler onresizeendEvent
		{
			add
			{
				CreateEventBridge();
				_onresizeendEvent += value;
			}
			remove
			{
				_onresizeendEvent -= value;
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML, 4
		/// </summary>
		private event HTMLFrameBase_onmouseenterEventHandler _onmouseenterEvent;

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public event HTMLFrameBase_onmouseenterEventHandler onmouseenterEvent
		{
			add
			{
				CreateEventBridge();
				_onmouseenterEvent += value;
			}
			remove
			{
				_onmouseenterEvent -= value;
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML, 4
		/// </summary>
		private event HTMLFrameBase_onmouseleaveEventHandler _onmouseleaveEvent;

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public event HTMLFrameBase_onmouseleaveEventHandler onmouseleaveEvent
		{
			add
			{
				CreateEventBridge();
				_onmouseleaveEvent += value;
			}
			remove
			{
				_onmouseleaveEvent -= value;
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML, 4
		/// </summary>
		private event HTMLFrameBase_onmousewheelEventHandler _onmousewheelEvent;

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public event HTMLFrameBase_onmousewheelEventHandler onmousewheelEvent
		{
			add
			{
				CreateEventBridge();
				_onmousewheelEvent += value;
			}
			remove
			{
				_onmousewheelEvent -= value;
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML, 4
		/// </summary>
		private event HTMLFrameBase_onactivateEventHandler _onactivateEvent;

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public event HTMLFrameBase_onactivateEventHandler onactivateEvent
		{
			add
			{
				CreateEventBridge();
				_onactivateEvent += value;
			}
			remove
			{
				_onactivateEvent -= value;
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML, 4
		/// </summary>
		private event HTMLFrameBase_ondeactivateEventHandler _ondeactivateEvent;

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public event HTMLFrameBase_ondeactivateEventHandler ondeactivateEvent
		{
			add
			{
				CreateEventBridge();
				_ondeactivateEvent += value;
			}
			remove
			{
				_ondeactivateEvent -= value;
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML, 4
		/// </summary>
		private event HTMLFrameBase_onfocusinEventHandler _onfocusinEvent;

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public event HTMLFrameBase_onfocusinEventHandler onfocusinEvent
		{
			add
			{
				CreateEventBridge();
				_onfocusinEvent += value;
			}
			remove
			{
				_onfocusinEvent -= value;
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML, 4
		/// </summary>
		private event HTMLFrameBase_onfocusoutEventHandler _onfocusoutEvent;

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public event HTMLFrameBase_onfocusoutEventHandler onfocusoutEvent
		{
			add
			{
				CreateEventBridge();
				_onfocusoutEvent += value;
			}
			remove
			{
				_onfocusoutEvent -= value;
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
				_activeSinkId = SinkHelper.GetConnectionPoint(this, ref _connectPoint, HTMLControlElementEvents_SinkHelper.Id);


			if(HTMLControlElementEvents_SinkHelper.Id.Equals(_activeSinkId, StringComparison.InvariantCultureIgnoreCase))
			{
				_hTMLControlElementEvents_SinkHelper = new HTMLControlElementEvents_SinkHelper(this, _connectPoint);
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
			if( null != _hTMLControlElementEvents_SinkHelper)
			{
				_hTMLControlElementEvents_SinkHelper.Dispose();
				_hTMLControlElementEvents_SinkHelper = null;
			}

			_connectPoint = null;
		}
        
        #endregion

		#pragma warning restore
	}
}