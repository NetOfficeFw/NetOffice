using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice;
using NetOffice.Misc;

namespace NetOffice.MSHTMLApi
{

	#region Delegates

	#pragma warning disable
	public delegate void HTMLInputElement_onhelpEventHandler();
	public delegate void HTMLInputElement_onclickEventHandler();
	public delegate void HTMLInputElement_ondblclickEventHandler();
	public delegate void HTMLInputElement_onkeypressEventHandler();
	public delegate void HTMLInputElement_onkeydownEventHandler();
	public delegate void HTMLInputElement_onkeyupEventHandler();
	public delegate void HTMLInputElement_onmouseoutEventHandler();
	public delegate void HTMLInputElement_onmouseoverEventHandler();
	public delegate void HTMLInputElement_onmousemoveEventHandler();
	public delegate void HTMLInputElement_onmousedownEventHandler();
	public delegate void HTMLInputElement_onmouseupEventHandler();
	public delegate void HTMLInputElement_onselectstartEventHandler();
	public delegate void HTMLInputElement_onfilterchangeEventHandler();
	public delegate void HTMLInputElement_ondragstartEventHandler();
	public delegate void HTMLInputElement_onbeforeupdateEventHandler();
	public delegate void HTMLInputElement_onafterupdateEventHandler();
	public delegate void HTMLInputElement_onerrorupdateEventHandler();
	public delegate void HTMLInputElement_onrowexitEventHandler();
	public delegate void HTMLInputElement_onrowenterEventHandler();
	public delegate void HTMLInputElement_ondatasetchangedEventHandler();
	public delegate void HTMLInputElement_ondataavailableEventHandler();
	public delegate void HTMLInputElement_ondatasetcompleteEventHandler();
	public delegate void HTMLInputElement_onlosecaptureEventHandler();
	public delegate void HTMLInputElement_onpropertychangeEventHandler();
	public delegate void HTMLInputElement_onscrollEventHandler();
	public delegate void HTMLInputElement_onfocusEventHandler();
	public delegate void HTMLInputElement_onblurEventHandler();
	public delegate void HTMLInputElement_onresizeEventHandler();
	public delegate void HTMLInputElement_ondragEventHandler();
	public delegate void HTMLInputElement_ondragendEventHandler();
	public delegate void HTMLInputElement_ondragenterEventHandler();
	public delegate void HTMLInputElement_ondragoverEventHandler();
	public delegate void HTMLInputElement_ondragleaveEventHandler();
	public delegate void HTMLInputElement_ondropEventHandler();
	public delegate void HTMLInputElement_onbeforecutEventHandler();
	public delegate void HTMLInputElement_oncutEventHandler();
	public delegate void HTMLInputElement_onbeforecopyEventHandler();
	public delegate void HTMLInputElement_oncopyEventHandler();
	public delegate void HTMLInputElement_onbeforepasteEventHandler();
	public delegate void HTMLInputElement_onpasteEventHandler();
	public delegate void HTMLInputElement_oncontextmenuEventHandler();
	public delegate void HTMLInputElement_onrowsdeleteEventHandler();
	public delegate void HTMLInputElement_onrowsinsertedEventHandler();
	public delegate void HTMLInputElement_oncellchangeEventHandler();
	public delegate void HTMLInputElement_onreadystatechangeEventHandler();
	public delegate void HTMLInputElement_onbeforeeditfocusEventHandler();
	public delegate void HTMLInputElement_onlayoutcompleteEventHandler();
	public delegate void HTMLInputElement_onpageEventHandler();
	public delegate void HTMLInputElement_onbeforedeactivateEventHandler();
	public delegate void HTMLInputElement_onbeforeactivateEventHandler();
	public delegate void HTMLInputElement_onmoveEventHandler();
	public delegate void HTMLInputElement_oncontrolselectEventHandler();
	public delegate void HTMLInputElement_onmovestartEventHandler();
	public delegate void HTMLInputElement_onmoveendEventHandler();
	public delegate void HTMLInputElement_onresizestartEventHandler();
	public delegate void HTMLInputElement_onresizeendEventHandler();
	public delegate void HTMLInputElement_onmouseenterEventHandler();
	public delegate void HTMLInputElement_onmouseleaveEventHandler();
	public delegate void HTMLInputElement_onmousewheelEventHandler();
	public delegate void HTMLInputElement_onactivateEventHandler();
	public delegate void HTMLInputElement_ondeactivateEventHandler();
	public delegate void HTMLInputElement_onfocusinEventHandler();
	public delegate void HTMLInputElement_onfocusoutEventHandler();
	public delegate void HTMLInputElement_onchangeEventHandler();
	public delegate void HTMLInputElement_onselectEventHandler();
	public delegate void HTMLInputElement_onloadEventHandler();
	public delegate void HTMLInputElement_onerrorEventHandler();
	public delegate void HTMLInputElement_onabortEventHandler();
	#pragma warning restore

	#endregion

	///<summary>
	/// CoClass HTMLInputElement 
	/// SupportByVersion MSHTML, 4
	///</summary>
	[SupportByVersionAttribute("MSHTML", 4)]
	[EntityTypeAttribute(EntityType.IsCoClass)]
	public class HTMLInputElement : DispHTMLInputElement,IEventBinding
	{
		#pragma warning disable
		#region Fields
		
		private NetRuntimeSystem.Runtime.InteropServices.ComTypes.IConnectionPoint _connectPoint;
		private string _activeSinkId;
		private NetRuntimeSystem.Type _thisType;
		HTMLInputTextElementEvents_SinkHelper _hTMLInputTextElementEvents_SinkHelper;
	
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
                    _type = typeof(HTMLInputElement);
                return _type;
            }
        }
        
        #endregion
        		
		#region Construction

		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
		public HTMLInputElement(Core factory, ICOMObject parentObject, object comProxy) : base(factory, parentObject, comProxy)
		{
			
		}

        ///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
		public HTMLInputElement(ICOMObject parentObject, object comProxy) : base(parentObject, comProxy)
		{
			
		}

		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public HTMLInputElement(Core factory, ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(factory, parentObject, comProxy, comProxyType)
		{
			
		}

		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public HTMLInputElement(ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(parentObject, comProxy, comProxyType)
		{
			
		}
		
		///<param name="replacedObject">object to replaced. replacedObject are not usable after this action</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public HTMLInputElement(ICOMObject replacedObject) : base(replacedObject)
		{
			
		}
		
		///<summary>
        /// Creates a new instance of HTMLInputElement 
        ///</summary>		
		public HTMLInputElement():base("MSHTML.HTMLInputElement")
		{
			
		}
		
		///<summary>
        /// Creates a new instance of HTMLInputElement
        ///</summary>
        ///<param name="progId">registered ProgID</param>
		public HTMLInputElement(string progId):base(progId)
		{
			
		}

		#endregion

		#region Static CoClass Methods

		/// <summary>
        /// Returns all running MSHTML.HTMLInputElement objects from the environment/system
        /// </summary>
        /// <returns>an MSHTML.HTMLInputElement array</returns>
		public static NetOffice.MSHTMLApi.HTMLInputElement[] GetActiveInstances()
		{		
			IDisposableEnumeration proxyList = NetOffice.ProxyService.GetActiveInstances("MSHTML","HTMLInputElement");
			NetRuntimeSystem.Collections.Generic.List<NetOffice.MSHTMLApi.HTMLInputElement> resultList = new NetRuntimeSystem.Collections.Generic.List<NetOffice.MSHTMLApi.HTMLInputElement>();
			foreach(object proxy in proxyList)
				resultList.Add( new NetOffice.MSHTMLApi.HTMLInputElement(null, proxy) );
			return resultList.ToArray();
		}

		/// <summary>
        /// Returns a running MSHTML.HTMLInputElement object from the environment/system.
        /// </summary>
        /// <returns>an MSHTML.HTMLInputElement object or null</returns>
		public static NetOffice.MSHTMLApi.HTMLInputElement GetActiveInstance()
		{
			object proxy  = NetOffice.ProxyService.GetActiveInstance("MSHTML","HTMLInputElement", false);
			if(null != proxy)
				return new NetOffice.MSHTMLApi.HTMLInputElement(null, proxy);
			else
				return null;
		}

		/// <summary>
        /// Returns a running MSHTML.HTMLInputElement object from the environment/system. 
        /// </summary>
	    /// <param name="throwOnError">throw an exception if no object was found</param>
        /// <returns>an MSHTML.HTMLInputElement object or null</returns>
		public static NetOffice.MSHTMLApi.HTMLInputElement GetActiveInstance(bool throwOnError)
		{
			object proxy  = NetOffice.ProxyService.GetActiveInstance("MSHTML","HTMLInputElement", throwOnError);
			if(null != proxy)
				return new NetOffice.MSHTMLApi.HTMLInputElement(null, proxy);
			else
				return null;
		}
		#endregion

		#region Events

		/// <summary>
		/// SupportByVersion MSHTML, 4
		/// </summary>
		private event HTMLInputElement_onhelpEventHandler _onhelpEvent;

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public event HTMLInputElement_onhelpEventHandler onhelpEvent
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
		private event HTMLInputElement_onclickEventHandler _onclickEvent;

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public event HTMLInputElement_onclickEventHandler onclickEvent
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
		private event HTMLInputElement_ondblclickEventHandler _ondblclickEvent;

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public event HTMLInputElement_ondblclickEventHandler ondblclickEvent
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
		private event HTMLInputElement_onkeypressEventHandler _onkeypressEvent;

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public event HTMLInputElement_onkeypressEventHandler onkeypressEvent
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
		private event HTMLInputElement_onkeydownEventHandler _onkeydownEvent;

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public event HTMLInputElement_onkeydownEventHandler onkeydownEvent
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
		private event HTMLInputElement_onkeyupEventHandler _onkeyupEvent;

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public event HTMLInputElement_onkeyupEventHandler onkeyupEvent
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
		private event HTMLInputElement_onmouseoutEventHandler _onmouseoutEvent;

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public event HTMLInputElement_onmouseoutEventHandler onmouseoutEvent
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
		private event HTMLInputElement_onmouseoverEventHandler _onmouseoverEvent;

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public event HTMLInputElement_onmouseoverEventHandler onmouseoverEvent
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
		private event HTMLInputElement_onmousemoveEventHandler _onmousemoveEvent;

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public event HTMLInputElement_onmousemoveEventHandler onmousemoveEvent
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
		private event HTMLInputElement_onmousedownEventHandler _onmousedownEvent;

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public event HTMLInputElement_onmousedownEventHandler onmousedownEvent
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
		private event HTMLInputElement_onmouseupEventHandler _onmouseupEvent;

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public event HTMLInputElement_onmouseupEventHandler onmouseupEvent
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
		private event HTMLInputElement_onselectstartEventHandler _onselectstartEvent;

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public event HTMLInputElement_onselectstartEventHandler onselectstartEvent
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
		private event HTMLInputElement_onfilterchangeEventHandler _onfilterchangeEvent;

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public event HTMLInputElement_onfilterchangeEventHandler onfilterchangeEvent
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
		private event HTMLInputElement_ondragstartEventHandler _ondragstartEvent;

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public event HTMLInputElement_ondragstartEventHandler ondragstartEvent
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
		private event HTMLInputElement_onbeforeupdateEventHandler _onbeforeupdateEvent;

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public event HTMLInputElement_onbeforeupdateEventHandler onbeforeupdateEvent
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
		private event HTMLInputElement_onafterupdateEventHandler _onafterupdateEvent;

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public event HTMLInputElement_onafterupdateEventHandler onafterupdateEvent
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
		private event HTMLInputElement_onerrorupdateEventHandler _onerrorupdateEvent;

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public event HTMLInputElement_onerrorupdateEventHandler onerrorupdateEvent
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
		private event HTMLInputElement_onrowexitEventHandler _onrowexitEvent;

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public event HTMLInputElement_onrowexitEventHandler onrowexitEvent
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
		private event HTMLInputElement_onrowenterEventHandler _onrowenterEvent;

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public event HTMLInputElement_onrowenterEventHandler onrowenterEvent
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
		private event HTMLInputElement_ondatasetchangedEventHandler _ondatasetchangedEvent;

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public event HTMLInputElement_ondatasetchangedEventHandler ondatasetchangedEvent
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
		private event HTMLInputElement_ondataavailableEventHandler _ondataavailableEvent;

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public event HTMLInputElement_ondataavailableEventHandler ondataavailableEvent
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
		private event HTMLInputElement_ondatasetcompleteEventHandler _ondatasetcompleteEvent;

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public event HTMLInputElement_ondatasetcompleteEventHandler ondatasetcompleteEvent
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
		private event HTMLInputElement_onlosecaptureEventHandler _onlosecaptureEvent;

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public event HTMLInputElement_onlosecaptureEventHandler onlosecaptureEvent
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
		private event HTMLInputElement_onpropertychangeEventHandler _onpropertychangeEvent;

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public event HTMLInputElement_onpropertychangeEventHandler onpropertychangeEvent
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
		private event HTMLInputElement_onscrollEventHandler _onscrollEvent;

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public event HTMLInputElement_onscrollEventHandler onscrollEvent
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
		private event HTMLInputElement_onfocusEventHandler _onfocusEvent;

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public event HTMLInputElement_onfocusEventHandler onfocusEvent
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
		private event HTMLInputElement_onblurEventHandler _onblurEvent;

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public event HTMLInputElement_onblurEventHandler onblurEvent
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
		private event HTMLInputElement_onresizeEventHandler _onresizeEvent;

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public event HTMLInputElement_onresizeEventHandler onresizeEvent
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
		private event HTMLInputElement_ondragEventHandler _ondragEvent;

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public event HTMLInputElement_ondragEventHandler ondragEvent
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
		private event HTMLInputElement_ondragendEventHandler _ondragendEvent;

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public event HTMLInputElement_ondragendEventHandler ondragendEvent
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
		private event HTMLInputElement_ondragenterEventHandler _ondragenterEvent;

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public event HTMLInputElement_ondragenterEventHandler ondragenterEvent
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
		private event HTMLInputElement_ondragoverEventHandler _ondragoverEvent;

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public event HTMLInputElement_ondragoverEventHandler ondragoverEvent
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
		private event HTMLInputElement_ondragleaveEventHandler _ondragleaveEvent;

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public event HTMLInputElement_ondragleaveEventHandler ondragleaveEvent
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
		private event HTMLInputElement_ondropEventHandler _ondropEvent;

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public event HTMLInputElement_ondropEventHandler ondropEvent
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
		private event HTMLInputElement_onbeforecutEventHandler _onbeforecutEvent;

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public event HTMLInputElement_onbeforecutEventHandler onbeforecutEvent
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
		private event HTMLInputElement_oncutEventHandler _oncutEvent;

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public event HTMLInputElement_oncutEventHandler oncutEvent
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
		private event HTMLInputElement_onbeforecopyEventHandler _onbeforecopyEvent;

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public event HTMLInputElement_onbeforecopyEventHandler onbeforecopyEvent
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
		private event HTMLInputElement_oncopyEventHandler _oncopyEvent;

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public event HTMLInputElement_oncopyEventHandler oncopyEvent
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
		private event HTMLInputElement_onbeforepasteEventHandler _onbeforepasteEvent;

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public event HTMLInputElement_onbeforepasteEventHandler onbeforepasteEvent
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
		private event HTMLInputElement_onpasteEventHandler _onpasteEvent;

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public event HTMLInputElement_onpasteEventHandler onpasteEvent
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
		private event HTMLInputElement_oncontextmenuEventHandler _oncontextmenuEvent;

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public event HTMLInputElement_oncontextmenuEventHandler oncontextmenuEvent
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
		private event HTMLInputElement_onrowsdeleteEventHandler _onrowsdeleteEvent;

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public event HTMLInputElement_onrowsdeleteEventHandler onrowsdeleteEvent
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
		private event HTMLInputElement_onrowsinsertedEventHandler _onrowsinsertedEvent;

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public event HTMLInputElement_onrowsinsertedEventHandler onrowsinsertedEvent
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
		private event HTMLInputElement_oncellchangeEventHandler _oncellchangeEvent;

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public event HTMLInputElement_oncellchangeEventHandler oncellchangeEvent
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
		private event HTMLInputElement_onreadystatechangeEventHandler _onreadystatechangeEvent;

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public event HTMLInputElement_onreadystatechangeEventHandler onreadystatechangeEvent
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
		private event HTMLInputElement_onbeforeeditfocusEventHandler _onbeforeeditfocusEvent;

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public event HTMLInputElement_onbeforeeditfocusEventHandler onbeforeeditfocusEvent
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
		private event HTMLInputElement_onlayoutcompleteEventHandler _onlayoutcompleteEvent;

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public event HTMLInputElement_onlayoutcompleteEventHandler onlayoutcompleteEvent
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
		private event HTMLInputElement_onpageEventHandler _onpageEvent;

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public event HTMLInputElement_onpageEventHandler onpageEvent
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
		private event HTMLInputElement_onbeforedeactivateEventHandler _onbeforedeactivateEvent;

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public event HTMLInputElement_onbeforedeactivateEventHandler onbeforedeactivateEvent
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
		private event HTMLInputElement_onbeforeactivateEventHandler _onbeforeactivateEvent;

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public event HTMLInputElement_onbeforeactivateEventHandler onbeforeactivateEvent
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
		private event HTMLInputElement_onmoveEventHandler _onmoveEvent;

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public event HTMLInputElement_onmoveEventHandler onmoveEvent
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
		private event HTMLInputElement_oncontrolselectEventHandler _oncontrolselectEvent;

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public event HTMLInputElement_oncontrolselectEventHandler oncontrolselectEvent
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
		private event HTMLInputElement_onmovestartEventHandler _onmovestartEvent;

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public event HTMLInputElement_onmovestartEventHandler onmovestartEvent
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
		private event HTMLInputElement_onmoveendEventHandler _onmoveendEvent;

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public event HTMLInputElement_onmoveendEventHandler onmoveendEvent
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
		private event HTMLInputElement_onresizestartEventHandler _onresizestartEvent;

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public event HTMLInputElement_onresizestartEventHandler onresizestartEvent
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
		private event HTMLInputElement_onresizeendEventHandler _onresizeendEvent;

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public event HTMLInputElement_onresizeendEventHandler onresizeendEvent
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
		private event HTMLInputElement_onmouseenterEventHandler _onmouseenterEvent;

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public event HTMLInputElement_onmouseenterEventHandler onmouseenterEvent
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
		private event HTMLInputElement_onmouseleaveEventHandler _onmouseleaveEvent;

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public event HTMLInputElement_onmouseleaveEventHandler onmouseleaveEvent
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
		private event HTMLInputElement_onmousewheelEventHandler _onmousewheelEvent;

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public event HTMLInputElement_onmousewheelEventHandler onmousewheelEvent
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
		private event HTMLInputElement_onactivateEventHandler _onactivateEvent;

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public event HTMLInputElement_onactivateEventHandler onactivateEvent
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
		private event HTMLInputElement_ondeactivateEventHandler _ondeactivateEvent;

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public event HTMLInputElement_ondeactivateEventHandler ondeactivateEvent
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
		private event HTMLInputElement_onfocusinEventHandler _onfocusinEvent;

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public event HTMLInputElement_onfocusinEventHandler onfocusinEvent
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
		private event HTMLInputElement_onfocusoutEventHandler _onfocusoutEvent;

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public event HTMLInputElement_onfocusoutEventHandler onfocusoutEvent
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

		/// <summary>
		/// SupportByVersion MSHTML, 4
		/// </summary>
		private event HTMLInputElement_onchangeEventHandler _onchangeEvent;

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public event HTMLInputElement_onchangeEventHandler onchangeEvent
		{
			add
			{
				CreateEventBridge();
				_onchangeEvent += value;
			}
			remove
			{
				_onchangeEvent -= value;
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML, 4
		/// </summary>
		private event HTMLInputElement_onselectEventHandler _onselectEvent;

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public event HTMLInputElement_onselectEventHandler onselectEvent
		{
			add
			{
				CreateEventBridge();
				_onselectEvent += value;
			}
			remove
			{
				_onselectEvent -= value;
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML, 4
		/// </summary>
		private event HTMLInputElement_onloadEventHandler _onloadEvent;

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public event HTMLInputElement_onloadEventHandler onloadEvent
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
		private event HTMLInputElement_onerrorEventHandler _onerrorEvent;

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public event HTMLInputElement_onerrorEventHandler onerrorEvent
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
		private event HTMLInputElement_onabortEventHandler _onabortEvent;

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public event HTMLInputElement_onabortEventHandler onabortEvent
		{
			add
			{
				CreateEventBridge();
				_onabortEvent += value;
			}
			remove
			{
				_onabortEvent -= value;
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
				_activeSinkId = SinkHelper.GetConnectionPoint(this, ref _connectPoint, HTMLInputTextElementEvents_SinkHelper.Id);


			if(HTMLInputTextElementEvents_SinkHelper.Id.Equals(_activeSinkId, StringComparison.InvariantCultureIgnoreCase))
			{
				_hTMLInputTextElementEvents_SinkHelper = new HTMLInputTextElementEvents_SinkHelper(this, _connectPoint);
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
			if( null != _hTMLInputTextElementEvents_SinkHelper)
			{
				_hTMLInputTextElementEvents_SinkHelper.Dispose();
				_hTMLInputTextElementEvents_SinkHelper = null;
			}

			_connectPoint = null;
		}
        
        #endregion

		#pragma warning restore
	}
}