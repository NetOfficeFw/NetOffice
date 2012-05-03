using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice;
namespace NetOffice.OutlookApi
{

	#region Delegates

	#pragma warning disable
	public delegate void OlkCommandButton_ClickEventHandler();
	public delegate void OlkCommandButton_DoubleClickEventHandler();
	public delegate void OlkCommandButton_MouseDownEventHandler(NetOffice.OutlookApi.Enums.OlMouseButton Button, NetOffice.OutlookApi.Enums.OlShiftState Shift, Single X, Single Y);
	public delegate void OlkCommandButton_MouseMoveEventHandler(NetOffice.OutlookApi.Enums.OlMouseButton Button, NetOffice.OutlookApi.Enums.OlShiftState Shift, Single X, Single Y);
	public delegate void OlkCommandButton_MouseUpEventHandler(NetOffice.OutlookApi.Enums.OlMouseButton Button, NetOffice.OutlookApi.Enums.OlShiftState Shift, Single X, Single Y);
	public delegate void OlkCommandButton_EnterEventHandler();
	public delegate void OlkCommandButton_ExitEventHandler(ref bool Cancel);
	public delegate void OlkCommandButton_KeyDownEventHandler(ref Int32 KeyCode, NetOffice.OutlookApi.Enums.OlShiftState Shift);
	public delegate void OlkCommandButton_KeyPressEventHandler(ref Int32 KeyAscii);
	public delegate void OlkCommandButton_KeyUpEventHandler(ref Int32 KeyCode, NetOffice.OutlookApi.Enums.OlShiftState Shift);
	public delegate void OlkCommandButton_AfterUpdateEventHandler();
	public delegate void OlkCommandButton_BeforeUpdateEventHandler(ref bool Cancel);
	#pragma warning restore

	#endregion

	///<summary>
	/// CoClass OlkCommandButton 
	/// SupportByVersion Outlook, 12,14
	///</summary>
	[SupportByVersionAttribute("Outlook", 12,14)]
	[EntityTypeAttribute(EntityType.IsCoClass)]
	public class OlkCommandButton : _OlkCommandButton,IEventBinding
	{
		#pragma warning disable
		#region Fields
		
		private NetRuntimeSystem.Runtime.InteropServices.ComTypes.IConnectionPoint _connectPoint;
		private string _activeSinkId;
		private NetRuntimeSystem.Type _thisType;
		OlkCommandButtonEvents_SinkHelper _olkCommandButtonEvents_SinkHelper;
	
		#endregion

		#region Type Information

        private static Type _type;
		
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        public static Type LateBindingApiWrapperType
        {
            get
            {
                if (null == _type)
                    _type = typeof(OlkCommandButton);
                    
                return _type;
            }
        }
        
        #endregion
        		
		#region Construction

        /// <param name="parentObject">object there has created the proxy</param>
        /// <param name="comProxy">inner wrapped COM proxy</param>
		public OlkCommandButton(COMObject parentObject, object comProxy) : base(parentObject, comProxy)
		{
			
		}

		/// <param name="parentObject">object there has created the proxy</param>
        /// <param name="comProxy">inner wrapped COM proxy</param>
        /// <param name="comProxyType">Type of inner wrapped COM proxy"</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public OlkCommandButton(COMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(parentObject, comProxy, comProxyType)
		{
			
		}
		
		/// <param name="replacedObject">object to replaced. replacedObject are not usable after this action</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public OlkCommandButton(COMObject replacedObject) : base(replacedObject)
		{
			
		}
		
		/// <summary>
        /// creates a new instance of OlkCommandButton 
        /// </summary>		
		public OlkCommandButton():base("Outlook.OlkCommandButton")
		{
			
		}
		
		/// <summary>
        /// creates a new instance of OlkCommandButton
        /// </summary>
        /// <param name="progId">registered ProgID</param>
		public OlkCommandButton(string progId):base(progId)
		{
			
		}

		#endregion

		#region Events

		/// <summary>
		/// SupportByVersion Outlook, 12,14
		/// </summary>
		private event OlkCommandButton_ClickEventHandler _ClickEvent;

		/// <summary>
		/// SupportByVersion Outlook 12 14
		/// </summary>
		[SupportByVersion("Outlook", 12,14)]
		public event OlkCommandButton_ClickEventHandler ClickEvent
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
		/// SupportByVersion Outlook, 12,14
		/// </summary>
		private event OlkCommandButton_DoubleClickEventHandler _DoubleClickEvent;

		/// <summary>
		/// SupportByVersion Outlook 12 14
		/// </summary>
		[SupportByVersion("Outlook", 12,14)]
		public event OlkCommandButton_DoubleClickEventHandler DoubleClickEvent
		{
			add
			{
				CreateEventBridge();
				_DoubleClickEvent += value;
			}
			remove
			{
				_DoubleClickEvent -= value;
			}
		}

		/// <summary>
		/// SupportByVersion Outlook, 12,14
		/// </summary>
		private event OlkCommandButton_MouseDownEventHandler _MouseDownEvent;

		/// <summary>
		/// SupportByVersion Outlook 12 14
		/// </summary>
		[SupportByVersion("Outlook", 12,14)]
		public event OlkCommandButton_MouseDownEventHandler MouseDownEvent
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
		/// SupportByVersion Outlook, 12,14
		/// </summary>
		private event OlkCommandButton_MouseMoveEventHandler _MouseMoveEvent;

		/// <summary>
		/// SupportByVersion Outlook 12 14
		/// </summary>
		[SupportByVersion("Outlook", 12,14)]
		public event OlkCommandButton_MouseMoveEventHandler MouseMoveEvent
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
		/// SupportByVersion Outlook, 12,14
		/// </summary>
		private event OlkCommandButton_MouseUpEventHandler _MouseUpEvent;

		/// <summary>
		/// SupportByVersion Outlook 12 14
		/// </summary>
		[SupportByVersion("Outlook", 12,14)]
		public event OlkCommandButton_MouseUpEventHandler MouseUpEvent
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
		/// SupportByVersion Outlook, 12,14
		/// </summary>
		private event OlkCommandButton_EnterEventHandler _EnterEvent;

		/// <summary>
		/// SupportByVersion Outlook 12 14
		/// </summary>
		[SupportByVersion("Outlook", 12,14)]
		public event OlkCommandButton_EnterEventHandler EnterEvent
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
		/// SupportByVersion Outlook, 12,14
		/// </summary>
		private event OlkCommandButton_ExitEventHandler _ExitEvent;

		/// <summary>
		/// SupportByVersion Outlook 12 14
		/// </summary>
		[SupportByVersion("Outlook", 12,14)]
		public event OlkCommandButton_ExitEventHandler ExitEvent
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
		/// SupportByVersion Outlook, 12,14
		/// </summary>
		private event OlkCommandButton_KeyDownEventHandler _KeyDownEvent;

		/// <summary>
		/// SupportByVersion Outlook 12 14
		/// </summary>
		[SupportByVersion("Outlook", 12,14)]
		public event OlkCommandButton_KeyDownEventHandler KeyDownEvent
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
		/// SupportByVersion Outlook, 12,14
		/// </summary>
		private event OlkCommandButton_KeyPressEventHandler _KeyPressEvent;

		/// <summary>
		/// SupportByVersion Outlook 12 14
		/// </summary>
		[SupportByVersion("Outlook", 12,14)]
		public event OlkCommandButton_KeyPressEventHandler KeyPressEvent
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
		/// SupportByVersion Outlook, 12,14
		/// </summary>
		private event OlkCommandButton_KeyUpEventHandler _KeyUpEvent;

		/// <summary>
		/// SupportByVersion Outlook 12 14
		/// </summary>
		[SupportByVersion("Outlook", 12,14)]
		public event OlkCommandButton_KeyUpEventHandler KeyUpEvent
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
		/// SupportByVersion Outlook, 12,14
		/// </summary>
		private event OlkCommandButton_AfterUpdateEventHandler _AfterUpdateEvent;

		/// <summary>
		/// SupportByVersion Outlook 12 14
		/// </summary>
		[SupportByVersion("Outlook", 12,14)]
		public event OlkCommandButton_AfterUpdateEventHandler AfterUpdateEvent
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
		/// SupportByVersion Outlook, 12,14
		/// </summary>
		private event OlkCommandButton_BeforeUpdateEventHandler _BeforeUpdateEvent;

		/// <summary>
		/// SupportByVersion Outlook 12 14
		/// </summary>
		[SupportByVersion("Outlook", 12,14)]
		public event OlkCommandButton_BeforeUpdateEventHandler BeforeUpdateEvent
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

		#endregion
       
	    #region IEventBinding Member
        
		/// <summary>
        /// creates active sink helper
        /// </summary>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public void CreateEventBridge()
        {
			if(false == NetOffice.Settings.EnableEvents)
				return;
	
			if (null != _connectPoint)
				return;
	
            if (null == _activeSinkId)
				_activeSinkId = SinkHelper.GetConnectionPoint(this, ref _connectPoint, OlkCommandButtonEvents_SinkHelper.Id);


			if(OlkCommandButtonEvents_SinkHelper.Id.Equals(_activeSinkId, StringComparison.InvariantCultureIgnoreCase))
			{
				_olkCommandButtonEvents_SinkHelper = new OlkCommandButtonEvents_SinkHelper(this, _connectPoint);
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
                        DebugConsole.WriteException(exception);
                    }
                }
                return delegates.Length;
            }
            else
                return 0;
		}

        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        public void DisposeEventBridge()
        {
			if( null != _olkCommandButtonEvents_SinkHelper)
			{
				_olkCommandButtonEvents_SinkHelper.Dispose();
				_olkCommandButtonEvents_SinkHelper = null;
			}

			_connectPoint = null;
		}
        
        #endregion

		#pragma warning restore
	}
}