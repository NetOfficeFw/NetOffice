using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.AccessApi
{
	#region Delegates

	#pragma warning disable
	public delegate void WebBrowserControl_UpdatedEventHandler(ref Int16 code);
	public delegate void WebBrowserControl_BeforeUpdateEventHandler(ref Int16 cancel);
	public delegate void WebBrowserControl_AfterUpdateEventHandler();
	public delegate void WebBrowserControl_EnterEventHandler();
	public delegate void WebBrowserControl_ExitEventHandler(ref Int16 cancel);
	public delegate void WebBrowserControl_DirtyEventHandler(ref Int16 cancel);
	public delegate void WebBrowserControl_ChangeEventHandler();
	public delegate void WebBrowserControl_GotFocusEventHandler();
	public delegate void WebBrowserControl_LostFocusEventHandler();
	public delegate void WebBrowserControl_ClickEventHandler();
	public delegate void WebBrowserControl_DblClickEventHandler(ref Int16 cancel);
	public delegate void WebBrowserControl_MouseDownEventHandler(ref Int16 button, ref Int16 shift, ref Single x, ref Single y);
	public delegate void WebBrowserControl_MouseMoveEventHandler(ref Int16 button, ref Int16 shift, ref Single x, ref Single y);
	public delegate void WebBrowserControl_MouseUpEventHandler(ref Int16 button, ref Int16 shift, ref Single x, ref Single y);
	public delegate void WebBrowserControl_KeyDownEventHandler(ref Int16 keyCode, ref Int16 shift);
	public delegate void WebBrowserControl_KeyPressEventHandler(ref Int16 keyAscii);
	public delegate void WebBrowserControl_KeyUpEventHandler(ref Int16 keyCode, ref Int16 shift);
	public delegate void WebBrowserControl_BeforeNavigate2EventHandler(ICOMObject pDisp, ref object url, ref object flags, ref object targetFrameName, ref object postData, ref object headers, ref bool cancel);
	public delegate void WebBrowserControl_DocumentCompleteEventHandler(ICOMObject pDisp, ref object url);
	public delegate void WebBrowserControl_ProgressChangeEventHandler(Int32 progress, Int32 progressMax);
	public delegate void WebBrowserControl_NavigateErrorEventHandler(ICOMObject pDisp, ref object url, ref object targetFrameName, ref object satusCode, ref bool cancel);
	#pragma warning restore

	#endregion

	/// <summary>
	/// CoClass WebBrowserControl 
	/// SupportByVersion Access, 14,15,16
	/// </summary>
	/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff835067.aspx </remarks>
	[SupportByVersion("Access", 14,15,16)]
	[EntityType(EntityType.IsCoClass)]
    [EventSink(typeof(Events.DispWebBrowserControlEvents_SinkHelper))]
    [ComEventInterface(typeof(Events.DispWebBrowserControlEvents))]
    public class WebBrowserControl : _WebBrowserControl, IEventBinding
	{
		#pragma warning disable

		#region Fields
		
		private NetRuntimeSystem.Runtime.InteropServices.ComTypes.IConnectionPoint _connectPoint;
		private string _activeSinkId;
        private static Type _type;
        private Events.DispWebBrowserControlEvents_SinkHelper _dispWebBrowserControlEvents_SinkHelper;
	
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
                    _type = typeof(WebBrowserControl);
                return _type;
            }
        }
        
        #endregion
        		
		#region Construction

		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
		public WebBrowserControl(Core factory, ICOMObject parentObject, object comProxy) : base(factory, parentObject, comProxy)
		{
			
		}

        ///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
		public WebBrowserControl(ICOMObject parentObject, object comProxy) : base(parentObject, comProxy)
		{
			
		}

		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public WebBrowserControl(Core factory, ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(factory, parentObject, comProxy, comProxyType)
		{
			
		}

		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public WebBrowserControl(ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(parentObject, comProxy, comProxyType)
		{
			
		}
		
		///<param name="replacedObject">object to replaced. replacedObject are not usable after this action</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public WebBrowserControl(ICOMObject replacedObject) : base(replacedObject)
		{
			
		}
		
		/// <summary>
        /// Creates a new instance of WebBrowserControl 
        /// </summary>		
		public WebBrowserControl():base("Access.WebBrowserControl")
		{
			
		}
		
		/// <summary>
        /// Creates a new instance of WebBrowserControl
        /// </summary>
        ///<param name="progId">registered ProgID</param>
		public WebBrowserControl(string progId):base(progId)
		{
			
		}

		#endregion

		#region Static CoClass Methods
		#endregion

		#region Events

		/// <summary>
		/// SupportByVersion Access, 14,15,16
		/// </summary>
		private event WebBrowserControl_UpdatedEventHandler _UpdatedEvent;

		/// <summary>
		/// SupportByVersion Access 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff196764.aspx </remarks>
		[SupportByVersion("Access", 14,15,16)]
		public event WebBrowserControl_UpdatedEventHandler UpdatedEvent
		{
			add
			{
				CreateEventBridge();
				_UpdatedEvent += value;
			}
			remove
			{
				_UpdatedEvent -= value;
			}
		}

		/// <summary>
		/// SupportByVersion Access, 14,15,16
		/// </summary>
		private event WebBrowserControl_BeforeUpdateEventHandler _BeforeUpdateEvent;

		/// <summary>
		/// SupportByVersion Access 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff195884.aspx </remarks>
		[SupportByVersion("Access", 14,15,16)]
		public event WebBrowserControl_BeforeUpdateEventHandler BeforeUpdateEvent
		{
			add
			{
				CreateEventBridge();
				_BeforeUpdateEvent += value;
			}
			remove
			{
				_BeforeUpdateEvent -= value;
			}
		}

		/// <summary>
		/// SupportByVersion Access, 14,15,16
		/// </summary>
		private event WebBrowserControl_AfterUpdateEventHandler _AfterUpdateEvent;

		/// <summary>
		/// SupportByVersion Access 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff197400.aspx </remarks>
		[SupportByVersion("Access", 14,15,16)]
		public event WebBrowserControl_AfterUpdateEventHandler AfterUpdateEvent
		{
			add
			{
				CreateEventBridge();
				_AfterUpdateEvent += value;
			}
			remove
			{
				_AfterUpdateEvent -= value;
			}
		}

		/// <summary>
		/// SupportByVersion Access, 14,15,16
		/// </summary>
		private event WebBrowserControl_EnterEventHandler _EnterEvent;

		/// <summary>
		/// SupportByVersion Access 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff193153.aspx </remarks>
		[SupportByVersion("Access", 14,15,16)]
		public event WebBrowserControl_EnterEventHandler EnterEvent
		{
			add
			{
				CreateEventBridge();
				_EnterEvent += value;
			}
			remove
			{
				_EnterEvent -= value;
			}
		}

		/// <summary>
		/// SupportByVersion Access, 14,15,16
		/// </summary>
		private event WebBrowserControl_ExitEventHandler _ExitEvent;

		/// <summary>
		/// SupportByVersion Access 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff821106.aspx </remarks>
		[SupportByVersion("Access", 14,15,16)]
		public event WebBrowserControl_ExitEventHandler ExitEvent
		{
			add
			{
				CreateEventBridge();
				_ExitEvent += value;
			}
			remove
			{
				_ExitEvent -= value;
			}
		}

		/// <summary>
		/// SupportByVersion Access, 14,15,16
		/// </summary>
		private event WebBrowserControl_DirtyEventHandler _DirtyEvent;

		/// <summary>
		/// SupportByVersion Access 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff192440.aspx </remarks>
		[SupportByVersion("Access", 14,15,16)]
		public event WebBrowserControl_DirtyEventHandler DirtyEvent
		{
			add
			{
				CreateEventBridge();
				_DirtyEvent += value;
			}
			remove
			{
				_DirtyEvent -= value;
			}
		}

		/// <summary>
		/// SupportByVersion Access, 14,15,16
		/// </summary>
		private event WebBrowserControl_ChangeEventHandler _ChangeEvent;

		/// <summary>
		/// SupportByVersion Access 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff192510.aspx </remarks>
		[SupportByVersion("Access", 14,15,16)]
		public event WebBrowserControl_ChangeEventHandler ChangeEvent
		{
			add
			{
				CreateEventBridge();
				_ChangeEvent += value;
			}
			remove
			{
				_ChangeEvent -= value;
			}
		}

		/// <summary>
		/// SupportByVersion Access, 14,15,16
		/// </summary>
		private event WebBrowserControl_GotFocusEventHandler _GotFocusEvent;

		/// <summary>
		/// SupportByVersion Access 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff195783.aspx </remarks>
		[SupportByVersion("Access", 14,15,16)]
		public event WebBrowserControl_GotFocusEventHandler GotFocusEvent
		{
			add
			{
				CreateEventBridge();
				_GotFocusEvent += value;
			}
			remove
			{
				_GotFocusEvent -= value;
			}
		}

		/// <summary>
		/// SupportByVersion Access, 14,15,16
		/// </summary>
		private event WebBrowserControl_LostFocusEventHandler _LostFocusEvent;

		/// <summary>
		/// SupportByVersion Access 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff193588.aspx </remarks>
		[SupportByVersion("Access", 14,15,16)]
		public event WebBrowserControl_LostFocusEventHandler LostFocusEvent
		{
			add
			{
				CreateEventBridge();
				_LostFocusEvent += value;
			}
			remove
			{
				_LostFocusEvent -= value;
			}
		}

		/// <summary>
		/// SupportByVersion Access, 14,15,16
		/// </summary>
		private event WebBrowserControl_ClickEventHandler _ClickEvent;

		/// <summary>
		/// SupportByVersion Access 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff192861.aspx </remarks>
		[SupportByVersion("Access", 14,15,16)]
		public event WebBrowserControl_ClickEventHandler ClickEvent
		{
			add
			{
				CreateEventBridge();
				_ClickEvent += value;
			}
			remove
			{
				_ClickEvent -= value;
			}
		}

		/// <summary>
		/// SupportByVersion Access, 14,15,16
		/// </summary>
		private event WebBrowserControl_DblClickEventHandler _DblClickEvent;

		/// <summary>
		/// SupportByVersion Access 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff835690.aspx </remarks>
		[SupportByVersion("Access", 14,15,16)]
		public event WebBrowserControl_DblClickEventHandler DblClickEvent
		{
			add
			{
				CreateEventBridge();
				_DblClickEvent += value;
			}
			remove
			{
				_DblClickEvent -= value;
			}
		}

		/// <summary>
		/// SupportByVersion Access, 14,15,16
		/// </summary>
		private event WebBrowserControl_MouseDownEventHandler _MouseDownEvent;

		/// <summary>
		/// SupportByVersion Access 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff823017.aspx </remarks>
		[SupportByVersion("Access", 14,15,16)]
		public event WebBrowserControl_MouseDownEventHandler MouseDownEvent
		{
			add
			{
				CreateEventBridge();
				_MouseDownEvent += value;
			}
			remove
			{
				_MouseDownEvent -= value;
			}
		}

		/// <summary>
		/// SupportByVersion Access, 14,15,16
		/// </summary>
		private event WebBrowserControl_MouseMoveEventHandler _MouseMoveEvent;

		/// <summary>
		/// SupportByVersion Access 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff845665.aspx </remarks>
		[SupportByVersion("Access", 14,15,16)]
		public event WebBrowserControl_MouseMoveEventHandler MouseMoveEvent
		{
			add
			{
				CreateEventBridge();
				_MouseMoveEvent += value;
			}
			remove
			{
				_MouseMoveEvent -= value;
			}
		}

		/// <summary>
		/// SupportByVersion Access, 14,15,16
		/// </summary>
		private event WebBrowserControl_MouseUpEventHandler _MouseUpEvent;

		/// <summary>
		/// SupportByVersion Access 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff196763.aspx </remarks>
		[SupportByVersion("Access", 14,15,16)]
		public event WebBrowserControl_MouseUpEventHandler MouseUpEvent
		{
			add
			{
				CreateEventBridge();
				_MouseUpEvent += value;
			}
			remove
			{
				_MouseUpEvent -= value;
			}
		}

		/// <summary>
		/// SupportByVersion Access, 14,15,16
		/// </summary>
		private event WebBrowserControl_KeyDownEventHandler _KeyDownEvent;

		/// <summary>
		/// SupportByVersion Access 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff845359.aspx </remarks>
		[SupportByVersion("Access", 14,15,16)]
		public event WebBrowserControl_KeyDownEventHandler KeyDownEvent
		{
			add
			{
				CreateEventBridge();
				_KeyDownEvent += value;
			}
			remove
			{
				_KeyDownEvent -= value;
			}
		}

		/// <summary>
		/// SupportByVersion Access, 14,15,16
		/// </summary>
		private event WebBrowserControl_KeyPressEventHandler _KeyPressEvent;

		/// <summary>
		/// SupportByVersion Access 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff194971.aspx </remarks>
		[SupportByVersion("Access", 14,15,16)]
		public event WebBrowserControl_KeyPressEventHandler KeyPressEvent
		{
			add
			{
				CreateEventBridge();
				_KeyPressEvent += value;
			}
			remove
			{
				_KeyPressEvent -= value;
			}
		}

		/// <summary>
		/// SupportByVersion Access, 14,15,16
		/// </summary>
		private event WebBrowserControl_KeyUpEventHandler _KeyUpEvent;

		/// <summary>
		/// SupportByVersion Access 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff835380.aspx </remarks>
		[SupportByVersion("Access", 14,15,16)]
		public event WebBrowserControl_KeyUpEventHandler KeyUpEvent
		{
			add
			{
				CreateEventBridge();
				_KeyUpEvent += value;
			}
			remove
			{
				_KeyUpEvent -= value;
			}
		}

		/// <summary>
		/// SupportByVersion Access, 14,15,16
		/// </summary>
		private event WebBrowserControl_BeforeNavigate2EventHandler _BeforeNavigate2Event;

		/// <summary>
		/// SupportByVersion Access 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff196461.aspx </remarks>
		[SupportByVersion("Access", 14,15,16)]
		public event WebBrowserControl_BeforeNavigate2EventHandler BeforeNavigate2Event
		{
			add
			{
				CreateEventBridge();
				_BeforeNavigate2Event += value;
			}
			remove
			{
				_BeforeNavigate2Event -= value;
			}
		}

		/// <summary>
		/// SupportByVersion Access, 14,15,16
		/// </summary>
		private event WebBrowserControl_DocumentCompleteEventHandler _DocumentCompleteEvent;

		/// <summary>
		/// SupportByVersion Access 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff197343.aspx </remarks>
		[SupportByVersion("Access", 14,15,16)]
		public event WebBrowserControl_DocumentCompleteEventHandler DocumentCompleteEvent
		{
			add
			{
				CreateEventBridge();
				_DocumentCompleteEvent += value;
			}
			remove
			{
				_DocumentCompleteEvent -= value;
			}
		}

		/// <summary>
		/// SupportByVersion Access, 14,15,16
		/// </summary>
		private event WebBrowserControl_ProgressChangeEventHandler _ProgressChangeEvent;

		/// <summary>
		/// SupportByVersion Access 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff845660.aspx </remarks>
		[SupportByVersion("Access", 14,15,16)]
		public event WebBrowserControl_ProgressChangeEventHandler ProgressChangeEvent
		{
			add
			{
				CreateEventBridge();
				_ProgressChangeEvent += value;
			}
			remove
			{
				_ProgressChangeEvent -= value;
			}
		}

		/// <summary>
		/// SupportByVersion Access, 14,15,16
		/// </summary>
		private event WebBrowserControl_NavigateErrorEventHandler _NavigateErrorEvent;

		/// <summary>
		/// SupportByVersion Access 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff845715.aspx </remarks>
		[SupportByVersion("Access", 14,15,16)]
		public event WebBrowserControl_NavigateErrorEventHandler NavigateErrorEvent
		{
			add
			{
				CreateEventBridge();
				_NavigateErrorEvent += value;
			}
			remove
			{
				_NavigateErrorEvent -= value;
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
				_activeSinkId = SinkHelper.GetConnectionPoint(this, ref _connectPoint, Events.DispWebBrowserControlEvents_SinkHelper.Id);


			if(Events.DispWebBrowserControlEvents_SinkHelper.Id.Equals(_activeSinkId, StringComparison.InvariantCultureIgnoreCase))
			{
				_dispWebBrowserControlEvents_SinkHelper = new Events.DispWebBrowserControlEvents_SinkHelper(this, _connectPoint);
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
			if( null != _dispWebBrowserControlEvents_SinkHelper)
			{
				_dispWebBrowserControlEvents_SinkHelper.Dispose();
				_dispWebBrowserControlEvents_SinkHelper = null;
			}

			_connectPoint = null;
		}
        
        #endregion

		#pragma warning restore
	}
}

