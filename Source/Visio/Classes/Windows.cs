﻿using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.VisioApi
{
	#region Delegates

	#pragma warning disable
	public delegate void Windows_WindowOpenedEventHandler(NetOffice.VisioApi.IVWindow window);
	public delegate void Windows_SelectionChangedEventHandler(NetOffice.VisioApi.IVWindow window);
	public delegate void Windows_BeforeWindowClosedEventHandler(NetOffice.VisioApi.IVWindow window);
	public delegate void Windows_WindowActivatedEventHandler(NetOffice.VisioApi.IVWindow window);
	public delegate void Windows_BeforeWindowSelDeleteEventHandler(NetOffice.VisioApi.IVWindow window);
	public delegate void Windows_BeforeWindowPageTurnEventHandler(NetOffice.VisioApi.IVWindow window);
	public delegate void Windows_WindowTurnedToPageEventHandler(NetOffice.VisioApi.IVWindow window);
	public delegate void Windows_WindowChangedEventHandler(NetOffice.VisioApi.IVWindow window);
	public delegate void Windows_ViewChangedEventHandler(NetOffice.VisioApi.IVWindow window);
	public delegate void Windows_QueryCancelWindowCloseEventHandler(NetOffice.VisioApi.IVWindow window);
	public delegate void Windows_WindowCloseCanceledEventHandler(NetOffice.VisioApi.IVWindow window);
	public delegate void Windows_OnKeystrokeMessageForAddonEventHandler(NetOffice.VisioApi.IVMSGWrap msg);
	public delegate void Windows_MouseDownEventHandler(Int32 button, Int32 keyButtonState, Double x, Double y, ref bool cancelDefault);
	public delegate void Windows_MouseMoveEventHandler(Int32 button, Int32 keyButtonState, Double x, Double y, ref bool cancelDefault);
	public delegate void Windows_MouseUpEventHandler(Int32 button, Int32 keyButtonState, Double x, Double y, ref bool cancelDefault);
	public delegate void Windows_KeyDownEventHandler(Int32 keyCode, Int32 keyButtonState, ref bool cancelDefault);
	public delegate void Windows_KeyPressEventHandler(Int32 keyAscii, ref bool cancelDefault);
	public delegate void Windows_KeyUpEventHandler(Int32 keyCode, Int32 keyButtonState, ref bool cancelDefault);
	#pragma warning restore

	#endregion

	/// <summary>
	/// CoClass Windows 
	/// SupportByVersion Visio, 11,12,14,15,16
	/// </summary>
	/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/ff769453(v=office.14).aspx </remarks>
	[SupportByVersion("Visio", 11,12,14,15,16)]
	[EntityType(EntityType.IsCoClass)]
	[EventSink(typeof(Events.EWindows_SinkHelper))]
    [ComEventInterface(typeof(Events.EWindows))]
    public class Windows : IVWindows, IEventBinding
	{
		#pragma warning disable

		#region Fields
		
		private NetRuntimeSystem.Runtime.InteropServices.ComTypes.IConnectionPoint _connectPoint;
		private string _activeSinkId;
        private static Type _type;
        private Events.EWindows_SinkHelper _eWindows_SinkHelper;
	
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
                    _type = typeof(Windows);
                return _type;
            }
        }
        
        #endregion
        		
		#region Construction

		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
		public Windows(Core factory, ICOMObject parentObject, object comProxy) : base(factory, parentObject, comProxy)
		{
			
		}

        ///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
		public Windows(ICOMObject parentObject, object comProxy) : base(parentObject, comProxy)
		{
			
		}

		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Windows(Core factory, ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(factory, parentObject, comProxy, comProxyType)
		{
			
		}

		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Windows(ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(parentObject, comProxy, comProxyType)
		{
			
		}
		
		///<param name="replacedObject">object to replaced. replacedObject are not usable after this action</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Windows(ICOMObject replacedObject) : base(replacedObject)
		{
			
		}
		
		/// <summary>
        /// Creates a new instance of Windows 
        /// </summary>		
		public Windows():base("Visio.Windows")
		{
			
		}
		
		/// <summary>
        /// Creates a new instance of Windows
        /// </summary>
        ///<param name="progId">registered ProgID</param>
		public Windows(string progId):base(progId)
		{
			
		}

		#endregion

		#region Static CoClass Methods
		#endregion

		#region Events

		/// <summary>
		/// SupportByVersion Visio, 11,12,14,15,16
		/// </summary>
		private event Windows_WindowOpenedEventHandler _WindowOpenedEvent;

		/// <summary>
		/// SupportByVersion Visio 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/ff765893(v=office.14).aspx </remarks>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public event Windows_WindowOpenedEventHandler WindowOpenedEvent
		{
			add
			{
				CreateEventBridge();
				_WindowOpenedEvent += value;
			}
			remove
			{
				_WindowOpenedEvent -= value;
			}
		}

		/// <summary>
		/// SupportByVersion Visio, 11,12,14,15,16
		/// </summary>
		private event Windows_SelectionChangedEventHandler _SelectionChangedEvent;

		/// <summary>
		/// SupportByVersion Visio 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/ff765788(v=office.14).aspx </remarks>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public event Windows_SelectionChangedEventHandler SelectionChangedEvent
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
		private event Windows_BeforeWindowClosedEventHandler _BeforeWindowClosedEvent;

		/// <summary>
		/// SupportByVersion Visio 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/ff769123(v=office.14).aspx </remarks>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public event Windows_BeforeWindowClosedEventHandler BeforeWindowClosedEvent
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
		private event Windows_WindowActivatedEventHandler _WindowActivatedEvent;

		/// <summary>
		/// SupportByVersion Visio 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/ff766087(v=office.14).aspx </remarks>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public event Windows_WindowActivatedEventHandler WindowActivatedEvent
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
		private event Windows_BeforeWindowSelDeleteEventHandler _BeforeWindowSelDeleteEvent;

		/// <summary>
		/// SupportByVersion Visio 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/ff768566(v=office.14).aspx </remarks>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public event Windows_BeforeWindowSelDeleteEventHandler BeforeWindowSelDeleteEvent
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
		private event Windows_BeforeWindowPageTurnEventHandler _BeforeWindowPageTurnEvent;

		/// <summary>
		/// SupportByVersion Visio 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/ff768773(v=office.14).aspx </remarks>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public event Windows_BeforeWindowPageTurnEventHandler BeforeWindowPageTurnEvent
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
		private event Windows_WindowTurnedToPageEventHandler _WindowTurnedToPageEvent;

		/// <summary>
		/// SupportByVersion Visio 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/ff768369(v=office.14).aspx </remarks>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public event Windows_WindowTurnedToPageEventHandler WindowTurnedToPageEvent
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
		private event Windows_WindowChangedEventHandler _WindowChangedEvent;

		/// <summary>
		/// SupportByVersion Visio 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/ff765071(v=office.14).aspx </remarks>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public event Windows_WindowChangedEventHandler WindowChangedEvent
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
		private event Windows_ViewChangedEventHandler _ViewChangedEvent;

		/// <summary>
		/// SupportByVersion Visio 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/ff765224(v=office.14).aspx </remarks>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public event Windows_ViewChangedEventHandler ViewChangedEvent
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
		private event Windows_QueryCancelWindowCloseEventHandler _QueryCancelWindowCloseEvent;

		/// <summary>
		/// SupportByVersion Visio 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/ff768016(v=office.14).aspx </remarks>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public event Windows_QueryCancelWindowCloseEventHandler QueryCancelWindowCloseEvent
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
		private event Windows_WindowCloseCanceledEventHandler _WindowCloseCanceledEvent;

		/// <summary>
		/// SupportByVersion Visio 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/ff766052(v=office.14).aspx </remarks>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public event Windows_WindowCloseCanceledEventHandler WindowCloseCanceledEvent
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
		private event Windows_OnKeystrokeMessageForAddonEventHandler _OnKeystrokeMessageForAddonEvent;

		/// <summary>
		/// SupportByVersion Visio 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/ff766300(v=office.14).aspx </remarks>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public event Windows_OnKeystrokeMessageForAddonEventHandler OnKeystrokeMessageForAddonEvent
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
		private event Windows_MouseDownEventHandler _MouseDownEvent;

		/// <summary>
		/// SupportByVersion Visio 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/ff768824(v=office.14).aspx </remarks>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public event Windows_MouseDownEventHandler MouseDownEvent
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
		private event Windows_MouseMoveEventHandler _MouseMoveEvent;

		/// <summary>
		/// SupportByVersion Visio 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/ff766046(v=office.14).aspx </remarks>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public event Windows_MouseMoveEventHandler MouseMoveEvent
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
		private event Windows_MouseUpEventHandler _MouseUpEvent;

		/// <summary>
		/// SupportByVersion Visio 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/ff765457(v=office.14).aspx </remarks>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public event Windows_MouseUpEventHandler MouseUpEvent
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
		private event Windows_KeyDownEventHandler _KeyDownEvent;

		/// <summary>
		/// SupportByVersion Visio 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/ff766065(v=office.14).aspx </remarks>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public event Windows_KeyDownEventHandler KeyDownEvent
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
		private event Windows_KeyPressEventHandler _KeyPressEvent;

		/// <summary>
		/// SupportByVersion Visio 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/ff766041(v=office.14).aspx </remarks>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public event Windows_KeyPressEventHandler KeyPressEvent
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
		private event Windows_KeyUpEventHandler _KeyUpEvent;

		/// <summary>
		/// SupportByVersion Visio 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/ff765374(v=office.14).aspx </remarks>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public event Windows_KeyUpEventHandler KeyUpEvent
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
				_activeSinkId = SinkHelper.GetConnectionPoint(this, ref _connectPoint, Events.EWindows_SinkHelper.Id);


			if(Events.EWindows_SinkHelper.Id.Equals(_activeSinkId, StringComparison.InvariantCultureIgnoreCase))
			{
				_eWindows_SinkHelper = new Events.EWindows_SinkHelper(this, _connectPoint);
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
			if( null != _eWindows_SinkHelper)
			{
				_eWindows_SinkHelper.Dispose();
				_eWindows_SinkHelper = null;
			}

			_connectPoint = null;
		}
        
        #endregion

		#pragma warning restore
	}
}