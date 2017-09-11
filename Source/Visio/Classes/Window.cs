using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.VisioApi
{
	#region Delegates

	#pragma warning disable
	public delegate void Window_SelectionChangedEventHandler(NetOffice.VisioApi.IVWindow window);
	public delegate void Window_BeforeWindowClosedEventHandler(NetOffice.VisioApi.IVWindow window);
	public delegate void Window_WindowActivatedEventHandler(NetOffice.VisioApi.IVWindow window);
	public delegate void Window_BeforeWindowSelDeleteEventHandler(NetOffice.VisioApi.IVWindow window);
	public delegate void Window_BeforeWindowPageTurnEventHandler(NetOffice.VisioApi.IVWindow window);
	public delegate void Window_WindowTurnedToPageEventHandler(NetOffice.VisioApi.IVWindow window);
	public delegate void Window_WindowChangedEventHandler(NetOffice.VisioApi.IVWindow window);
	public delegate void Window_ViewChangedEventHandler(NetOffice.VisioApi.IVWindow window);
	public delegate void Window_QueryCancelWindowCloseEventHandler(NetOffice.VisioApi.IVWindow window);
	public delegate void Window_WindowCloseCanceledEventHandler(NetOffice.VisioApi.IVWindow window);
	public delegate void Window_OnKeystrokeMessageForAddonEventHandler(NetOffice.VisioApi.IVMSGWrap msg);
	public delegate void Window_MouseDownEventHandler(Int32 button, Int32 keyButtonState, Double x, Double y, ref bool cancelDefault);
	public delegate void Window_MouseMoveEventHandler(Int32 button, Int32 keyButtonState, Double x, Double y, ref bool cancelDefault);
	public delegate void Window_MouseUpEventHandler(Int32 button, Int32 keyButtonState, Double x, Double y, ref bool cancelDefault);
	public delegate void Window_KeyDownEventHandler(Int32 KeyCode, Int32 keyButtonState, ref bool cancelDefault);
	public delegate void Window_KeyPressEventHandler(Int32 keyAscii, ref bool cancelDefault);
	public delegate void Window_KeyUpEventHandler(Int32 keyCode, Int32 keyButtonState, ref bool cancelDefault);
	#pragma warning restore

	#endregion

	/// <summary>
	/// CoClass Window 
	/// SupportByVersion Visio, 11,12,14,15,16
	/// </summary>
	/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/ff769449(v=office.14).aspx </remarks>
	[SupportByVersion("Visio", 11,12,14,15,16)]
	[EntityType(EntityType.IsCoClass)]
	[EventSink(typeof(Events.EWindow_SinkHelper))]
    [ComEventInterface(typeof(Events.EWindow))]
    public class Window : IVWindow, IEventBinding
	{
		#pragma warning disable

		#region Fields
		
		private NetRuntimeSystem.Runtime.InteropServices.ComTypes.IConnectionPoint _connectPoint;
		private string _activeSinkId;
        private static Type _type;
        private Events.EWindow_SinkHelper _eWindow_SinkHelper;
	
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
                    _type = typeof(Window);
                return _type;
            }
        }
        
        #endregion
        		
		#region Construction

		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
		public Window(Core factory, ICOMObject parentObject, object comProxy) : base(factory, parentObject, comProxy)
		{
			
		}

        ///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
		public Window(ICOMObject parentObject, object comProxy) : base(parentObject, comProxy)
		{
			
		}

		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Window(Core factory, ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(factory, parentObject, comProxy, comProxyType)
		{
			
		}

		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Window(ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(parentObject, comProxy, comProxyType)
		{
			
		}
		
		///<param name="replacedObject">object to replaced. replacedObject are not usable after this action</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Window(ICOMObject replacedObject) : base(replacedObject)
		{
			
		}
		
		/// <summary>
        /// Creates a new instance of Window 
        /// </summary>		
		public Window():base("Visio.Window")
		{
			
		}
		
		/// <summary>
        /// Creates a new instance of Window
        /// </summary>
        ///<param name="progId">registered ProgID</param>
		public Window(string progId):base(progId)
		{
			
		}

		#endregion

		#region Static CoClass Methods
		#endregion

		#region Events

		/// <summary>
		/// SupportByVersion Visio, 11,12,14,15,16
		/// </summary>
		private event Window_SelectionChangedEventHandler _SelectionChangedEvent;

		/// <summary>
		/// SupportByVersion Visio 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/ff766358(v=office.14).aspx </remarks>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public event Window_SelectionChangedEventHandler SelectionChangedEvent
		{
			add
			{
				CreateEventBridge();
				_SelectionChangedEvent += value;
			}
			remove
			{
				_SelectionChangedEvent -= value;
			}
		}

		/// <summary>
		/// SupportByVersion Visio, 11,12,14,15,16
		/// </summary>
		private event Window_BeforeWindowClosedEventHandler _BeforeWindowClosedEvent;

		/// <summary>
		/// SupportByVersion Visio 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/ff766158(v=office.14).aspx </remarks>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public event Window_BeforeWindowClosedEventHandler BeforeWindowClosedEvent
		{
			add
			{
				CreateEventBridge();
				_BeforeWindowClosedEvent += value;
			}
			remove
			{
				_BeforeWindowClosedEvent -= value;
			}
		}

		/// <summary>
		/// SupportByVersion Visio, 11,12,14,15,16
		/// </summary>
		private event Window_WindowActivatedEventHandler _WindowActivatedEvent;

		/// <summary>
		/// SupportByVersion Visio 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/ff767395(v=office.14).aspx </remarks>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public event Window_WindowActivatedEventHandler WindowActivatedEvent
		{
			add
			{
				CreateEventBridge();
				_WindowActivatedEvent += value;
			}
			remove
			{
				_WindowActivatedEvent -= value;
			}
		}

		/// <summary>
		/// SupportByVersion Visio, 11,12,14,15,16
		/// </summary>
		private event Window_BeforeWindowSelDeleteEventHandler _BeforeWindowSelDeleteEvent;

		/// <summary>
		/// SupportByVersion Visio 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/ff766155(v=office.14).aspx </remarks>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public event Window_BeforeWindowSelDeleteEventHandler BeforeWindowSelDeleteEvent
		{
			add
			{
				CreateEventBridge();
				_BeforeWindowSelDeleteEvent += value;
			}
			remove
			{
				_BeforeWindowSelDeleteEvent -= value;
			}
		}

		/// <summary>
		/// SupportByVersion Visio, 11,12,14,15,16
		/// </summary>
		private event Window_BeforeWindowPageTurnEventHandler _BeforeWindowPageTurnEvent;

		/// <summary>
		/// SupportByVersion Visio 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/ff767156(v=office.14).aspx </remarks>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public event Window_BeforeWindowPageTurnEventHandler BeforeWindowPageTurnEvent
		{
			add
			{
				CreateEventBridge();
				_BeforeWindowPageTurnEvent += value;
			}
			remove
			{
				_BeforeWindowPageTurnEvent -= value;
			}
		}

		/// <summary>
		/// SupportByVersion Visio, 11,12,14,15,16
		/// </summary>
		private event Window_WindowTurnedToPageEventHandler _WindowTurnedToPageEvent;

		/// <summary>
		/// SupportByVersion Visio 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/ff768977(v=office.14).aspx </remarks>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public event Window_WindowTurnedToPageEventHandler WindowTurnedToPageEvent
		{
			add
			{
				CreateEventBridge();
				_WindowTurnedToPageEvent += value;
			}
			remove
			{
				_WindowTurnedToPageEvent -= value;
			}
		}

		/// <summary>
		/// SupportByVersion Visio, 11,12,14,15,16
		/// </summary>
		private event Window_WindowChangedEventHandler _WindowChangedEvent;

		/// <summary>
		/// SupportByVersion Visio 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/ff768918(v=office.14).aspx </remarks>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public event Window_WindowChangedEventHandler WindowChangedEvent
		{
			add
			{
				CreateEventBridge();
				_WindowChangedEvent += value;
			}
			remove
			{
				_WindowChangedEvent -= value;
			}
		}

		/// <summary>
		/// SupportByVersion Visio, 11,12,14,15,16
		/// </summary>
		private event Window_ViewChangedEventHandler _ViewChangedEvent;

		/// <summary>
		/// SupportByVersion Visio 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/ff767710(v=office.14).aspx </remarks>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public event Window_ViewChangedEventHandler ViewChangedEvent
		{
			add
			{
				CreateEventBridge();
				_ViewChangedEvent += value;
			}
			remove
			{
				_ViewChangedEvent -= value;
			}
		}

		/// <summary>
		/// SupportByVersion Visio, 11,12,14,15,16
		/// </summary>
		private event Window_QueryCancelWindowCloseEventHandler _QueryCancelWindowCloseEvent;

		/// <summary>
		/// SupportByVersion Visio 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/ff766114(v=office.14).aspx </remarks>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public event Window_QueryCancelWindowCloseEventHandler QueryCancelWindowCloseEvent
		{
			add
			{
				CreateEventBridge();
				_QueryCancelWindowCloseEvent += value;
			}
			remove
			{
				_QueryCancelWindowCloseEvent -= value;
			}
		}

		/// <summary>
		/// SupportByVersion Visio, 11,12,14,15,16
		/// </summary>
		private event Window_WindowCloseCanceledEventHandler _WindowCloseCanceledEvent;

		/// <summary>
		/// SupportByVersion Visio 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/ff768122(v=office.14).aspx </remarks>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public event Window_WindowCloseCanceledEventHandler WindowCloseCanceledEvent
		{
			add
			{
				CreateEventBridge();
				_WindowCloseCanceledEvent += value;
			}
			remove
			{
				_WindowCloseCanceledEvent -= value;
			}
		}

		/// <summary>
		/// SupportByVersion Visio, 11,12,14,15,16
		/// </summary>
		private event Window_OnKeystrokeMessageForAddonEventHandler _OnKeystrokeMessageForAddonEvent;

		/// <summary>
		/// SupportByVersion Visio 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/ff767265(v=office.14).aspx </remarks>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public event Window_OnKeystrokeMessageForAddonEventHandler OnKeystrokeMessageForAddonEvent
		{
			add
			{
				CreateEventBridge();
				_OnKeystrokeMessageForAddonEvent += value;
			}
			remove
			{
				_OnKeystrokeMessageForAddonEvent -= value;
			}
		}

		/// <summary>
		/// SupportByVersion Visio, 11,12,14,15,16
		/// </summary>
		private event Window_MouseDownEventHandler _MouseDownEvent;

		/// <summary>
		/// SupportByVersion Visio 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/ff767550(v=office.14).aspx </remarks>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public event Window_MouseDownEventHandler MouseDownEvent
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
		/// SupportByVersion Visio, 11,12,14,15,16
		/// </summary>
		private event Window_MouseMoveEventHandler _MouseMoveEvent;

		/// <summary>
		/// SupportByVersion Visio 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/ff767508(v=office.14).aspx </remarks>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public event Window_MouseMoveEventHandler MouseMoveEvent
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
		/// SupportByVersion Visio, 11,12,14,15,16
		/// </summary>
		private event Window_MouseUpEventHandler _MouseUpEvent;

		/// <summary>
		/// SupportByVersion Visio 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/ff768316(v=office.14).aspx </remarks>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public event Window_MouseUpEventHandler MouseUpEvent
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
		/// SupportByVersion Visio, 11,12,14,15,16
		/// </summary>
		private event Window_KeyDownEventHandler _KeyDownEvent;

		/// <summary>
		/// SupportByVersion Visio 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/ff766901(v=office.14).aspx </remarks>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public event Window_KeyDownEventHandler KeyDownEvent
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
		/// SupportByVersion Visio, 11,12,14,15,16
		/// </summary>
		private event Window_KeyPressEventHandler _KeyPressEvent;

		/// <summary>
		/// SupportByVersion Visio 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/ff767366(v=office.14).aspx </remarks>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public event Window_KeyPressEventHandler KeyPressEvent
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
		/// SupportByVersion Visio, 11,12,14,15,16
		/// </summary>
		private event Window_KeyUpEventHandler _KeyUpEvent;

		/// <summary>
		/// SupportByVersion Visio 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/ff767880(v=office.14).aspx </remarks>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public event Window_KeyUpEventHandler KeyUpEvent
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
				_activeSinkId = SinkHelper.GetConnectionPoint(this, ref _connectPoint, Events.EWindow_SinkHelper.Id);


			if(Events.EWindow_SinkHelper.Id.Equals(_activeSinkId, StringComparison.InvariantCultureIgnoreCase))
			{
				_eWindow_SinkHelper = new Events.EWindow_SinkHelper(this, _connectPoint);
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
			if( null != _eWindow_SinkHelper)
			{
				_eWindow_SinkHelper.Dispose();
				_eWindow_SinkHelper = null;
			}

			_connectPoint = null;
		}
        
        #endregion

		#pragma warning restore
	}
}

