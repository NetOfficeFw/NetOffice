using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice;
namespace NetOffice.AccessApi
{

	#region Delegates

	#pragma warning disable
	public delegate void _ToggleButtonInOption_GotFocusEventHandler();
	public delegate void _ToggleButtonInOption_LostFocusEventHandler();
	public delegate void _ToggleButtonInOption_MouseDownEventHandler(ref Int16 Button, ref Int16 Shift, ref Single X, ref Single Y);
	public delegate void _ToggleButtonInOption_MouseMoveEventHandler(ref Int16 Button, ref Int16 Shift, ref Single X, ref Single Y);
	public delegate void _ToggleButtonInOption_MouseUpEventHandler(ref Int16 Button, ref Int16 Shift, ref Single X, ref Single Y);
	public delegate void _ToggleButtonInOption_KeyDownEventHandler(ref Int16 KeyCode, ref Int16 Shift);
	public delegate void _ToggleButtonInOption_KeyPressEventHandler(ref Int16 KeyAscii);
	public delegate void _ToggleButtonInOption_KeyUpEventHandler(ref Int16 KeyCode, ref Int16 Shift);
	public delegate void _ToggleButtonInOption_ClickEventHandler();
	public delegate void _ToggleButtonInOption_BeforeUpdateEventHandler(ref Int16 Cancel);
	public delegate void _ToggleButtonInOption_AfterUpdateEventHandler();
	public delegate void _ToggleButtonInOption_EnterEventHandler();
	public delegate void _ToggleButtonInOption_ExitEventHandler(ref Int16 Cancel);
	public delegate void _ToggleButtonInOption_DblClickEventHandler(ref Int16 Cancel);
	#pragma warning restore

	#endregion

	///<summary>
	/// CoClass _ToggleButtonInOption 
	/// SupportByVersion Access, 9,10,11,12,14,15
	///</summary>
	[SupportByVersionAttribute("Access", 9,10,11,12,14,15)]
	[EntityTypeAttribute(EntityType.IsCoClass)]
	public class _ToggleButtonInOption : _ToggleButton,IEventBinding
	{
		#pragma warning disable
		#region Fields
		
		private NetRuntimeSystem.Runtime.InteropServices.ComTypes.IConnectionPoint _connectPoint;
		private string _activeSinkId;
		private NetRuntimeSystem.Type _thisType;
		_ToggleButtonInOptionEvents_SinkHelper __ToggleButtonInOptionEvents_SinkHelper;
		DispToggleButtonEvents_SinkHelper _dispToggleButtonEvents_SinkHelper;
	
		#endregion

		#region Type Information

        private static Type _type;
		
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        public static Type LateBindingApiWrapperType
        {
            get
            {
                if (null == _type)
                    _type = typeof(_ToggleButtonInOption);
                return _type;
            }
        }
        
        #endregion
        		
		#region Construction

		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
		public _ToggleButtonInOption(Core factory, COMObject parentObject, object comProxy) : base(factory, parentObject, comProxy)
		{
			
		}

        ///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
		public _ToggleButtonInOption(COMObject parentObject, object comProxy) : base(parentObject, comProxy)
		{
			
		}

		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public _ToggleButtonInOption(Core factory, COMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(factory, parentObject, comProxy, comProxyType)
		{
			
		}

		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public _ToggleButtonInOption(COMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(parentObject, comProxy, comProxyType)
		{
			
		}
		
		///<param name="replacedObject">object to replaced. replacedObject are not usable after this action</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public _ToggleButtonInOption(COMObject replacedObject) : base(replacedObject)
		{
			
		}
		
		///<summary>
        ///creates a new instance of _ToggleButtonInOption 
        ///</summary>		
		public _ToggleButtonInOption():base("Access._ToggleButtonInOption")
		{
			
		}
		
		///<summary>
        ///creates a new instance of _ToggleButtonInOption
        ///</summary>
        ///<param name="progId">registered ProgID</param>
		public _ToggleButtonInOption(string progId):base(progId)
		{
			
		}

		#endregion

		#region Static CoClass Methods

		/// <summary>
        /// returns all running Access._ToggleButtonInOption objects from the running object table(ROT)
        /// </summary>
        /// <returns>an Access._ToggleButtonInOption array</returns>
		public static NetOffice.AccessApi._ToggleButtonInOption[] GetActiveInstances()
		{		
			NetRuntimeSystem.Collections.Generic.List<object> proxyList = NetOffice.RunningObjectTable.GetActiveProxiesFromROT("Access","_ToggleButtonInOption");
			NetRuntimeSystem.Collections.Generic.List<NetOffice.AccessApi._ToggleButtonInOption> resultList = new NetRuntimeSystem.Collections.Generic.List<NetOffice.AccessApi._ToggleButtonInOption>();
			foreach(object proxy in proxyList)
				resultList.Add( new NetOffice.AccessApi._ToggleButtonInOption(null, proxy) );
			return resultList.ToArray();
		}

		/// <summary>
        /// returns a running Access._ToggleButtonInOption object from the running object table(ROT). the method takes the first element from the table
        /// </summary>
        /// <returns>an Access._ToggleButtonInOption object or null</returns>
		public static NetOffice.AccessApi._ToggleButtonInOption GetActiveInstance()
		{
			object proxy = NetOffice.RunningObjectTable.GetActiveProxyFromROT("Access","_ToggleButtonInOption", false);
			if(null != proxy)
				return new NetOffice.AccessApi._ToggleButtonInOption(null, proxy);
			else
				return null;
		}

		/// <summary>
        /// returns a running Access._ToggleButtonInOption object from the running object table(ROT).  the method takes the first element from the table
        /// </summary>
	    /// <param name="throwOnError">throw an exception if no object was found</param>
        /// <returns>an Access._ToggleButtonInOption object or null</returns>
		public static NetOffice.AccessApi._ToggleButtonInOption GetActiveInstance(bool throwOnError)
		{
			object proxy = NetOffice.RunningObjectTable.GetActiveProxyFromROT("Access","_ToggleButtonInOption", throwOnError);
			if(null != proxy)
				return new NetOffice.AccessApi._ToggleButtonInOption(null, proxy);
			else
				return null;
		}
		#endregion

		#region Events

		/// <summary>
		/// SupportByVersion Access, 9,10,11,12,14,15
		/// </summary>
		private event _ToggleButtonInOption_GotFocusEventHandler _GotFocusEvent;

		/// <summary>
		/// SupportByVersion Access 9 10 11 12 14 15
		/// </summary>
		[SupportByVersion("Access", 9,10,11,12,14,15)]
		public event _ToggleButtonInOption_GotFocusEventHandler GotFocusEvent
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
		/// SupportByVersion Access, 9,10,11,12,14,15
		/// </summary>
		private event _ToggleButtonInOption_LostFocusEventHandler _LostFocusEvent;

		/// <summary>
		/// SupportByVersion Access 9 10 11 12 14 15
		/// </summary>
		[SupportByVersion("Access", 9,10,11,12,14,15)]
		public event _ToggleButtonInOption_LostFocusEventHandler LostFocusEvent
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
		/// SupportByVersion Access, 9,10,11,12,14,15
		/// </summary>
		private event _ToggleButtonInOption_MouseDownEventHandler _MouseDownEvent;

		/// <summary>
		/// SupportByVersion Access 9 10 11 12 14 15
		/// </summary>
		[SupportByVersion("Access", 9,10,11,12,14,15)]
		public event _ToggleButtonInOption_MouseDownEventHandler MouseDownEvent
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
		/// SupportByVersion Access, 9,10,11,12,14,15
		/// </summary>
		private event _ToggleButtonInOption_MouseMoveEventHandler _MouseMoveEvent;

		/// <summary>
		/// SupportByVersion Access 9 10 11 12 14 15
		/// </summary>
		[SupportByVersion("Access", 9,10,11,12,14,15)]
		public event _ToggleButtonInOption_MouseMoveEventHandler MouseMoveEvent
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
		/// SupportByVersion Access, 9,10,11,12,14,15
		/// </summary>
		private event _ToggleButtonInOption_MouseUpEventHandler _MouseUpEvent;

		/// <summary>
		/// SupportByVersion Access 9 10 11 12 14 15
		/// </summary>
		[SupportByVersion("Access", 9,10,11,12,14,15)]
		public event _ToggleButtonInOption_MouseUpEventHandler MouseUpEvent
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
		/// SupportByVersion Access, 9,10,11,12,14,15
		/// </summary>
		private event _ToggleButtonInOption_KeyDownEventHandler _KeyDownEvent;

		/// <summary>
		/// SupportByVersion Access 9 10 11 12 14 15
		/// </summary>
		[SupportByVersion("Access", 9,10,11,12,14,15)]
		public event _ToggleButtonInOption_KeyDownEventHandler KeyDownEvent
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
		/// SupportByVersion Access, 9,10,11,12,14,15
		/// </summary>
		private event _ToggleButtonInOption_KeyPressEventHandler _KeyPressEvent;

		/// <summary>
		/// SupportByVersion Access 9 10 11 12 14 15
		/// </summary>
		[SupportByVersion("Access", 9,10,11,12,14,15)]
		public event _ToggleButtonInOption_KeyPressEventHandler KeyPressEvent
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
		/// SupportByVersion Access, 9,10,11,12,14,15
		/// </summary>
		private event _ToggleButtonInOption_KeyUpEventHandler _KeyUpEvent;

		/// <summary>
		/// SupportByVersion Access 9 10 11 12 14 15
		/// </summary>
		[SupportByVersion("Access", 9,10,11,12,14,15)]
		public event _ToggleButtonInOption_KeyUpEventHandler KeyUpEvent
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
		/// SupportByVersion Access, 12,14,15
		/// </summary>
		private event _ToggleButtonInOption_ClickEventHandler _ClickEvent;

		/// <summary>
		/// SupportByVersion Access 12 14 15
		/// </summary>
		[SupportByVersion("Access", 12,14,15)]
		public event _ToggleButtonInOption_ClickEventHandler ClickEvent
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
		/// SupportByVersion Access, 12,14,15
		/// </summary>
		private event _ToggleButtonInOption_BeforeUpdateEventHandler _BeforeUpdateEvent;

		/// <summary>
		/// SupportByVersion Access 12 14 15
		/// </summary>
		[SupportByVersion("Access", 12,14,15)]
		public event _ToggleButtonInOption_BeforeUpdateEventHandler BeforeUpdateEvent
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
		/// SupportByVersion Access, 12,14,15
		/// </summary>
		private event _ToggleButtonInOption_AfterUpdateEventHandler _AfterUpdateEvent;

		/// <summary>
		/// SupportByVersion Access 12 14 15
		/// </summary>
		[SupportByVersion("Access", 12,14,15)]
		public event _ToggleButtonInOption_AfterUpdateEventHandler AfterUpdateEvent
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
		/// SupportByVersion Access, 12,14,15
		/// </summary>
		private event _ToggleButtonInOption_EnterEventHandler _EnterEvent;

		/// <summary>
		/// SupportByVersion Access 12 14 15
		/// </summary>
		[SupportByVersion("Access", 12,14,15)]
		public event _ToggleButtonInOption_EnterEventHandler EnterEvent
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
		/// SupportByVersion Access, 12,14,15
		/// </summary>
		private event _ToggleButtonInOption_ExitEventHandler _ExitEvent;

		/// <summary>
		/// SupportByVersion Access 12 14 15
		/// </summary>
		[SupportByVersion("Access", 12,14,15)]
		public event _ToggleButtonInOption_ExitEventHandler ExitEvent
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
		/// SupportByVersion Access, 12,14,15
		/// </summary>
		private event _ToggleButtonInOption_DblClickEventHandler _DblClickEvent;

		/// <summary>
		/// SupportByVersion Access 12 14 15
		/// </summary>
		[SupportByVersion("Access", 12,14,15)]
		public event _ToggleButtonInOption_DblClickEventHandler DblClickEvent
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

		#endregion
       
	    #region IEventBinding Member
        
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
				_activeSinkId = SinkHelper.GetConnectionPoint(this, ref _connectPoint, _ToggleButtonInOptionEvents_SinkHelper.Id,DispToggleButtonEvents_SinkHelper.Id);


			if(_ToggleButtonInOptionEvents_SinkHelper.Id.Equals(_activeSinkId, StringComparison.InvariantCultureIgnoreCase))
			{
				__ToggleButtonInOptionEvents_SinkHelper = new _ToggleButtonInOptionEvents_SinkHelper(this, _connectPoint);
				return;
			}

			if(DispToggleButtonEvents_SinkHelper.Id.Equals(_activeSinkId, StringComparison.InvariantCultureIgnoreCase))
			{
				_dispToggleButtonEvents_SinkHelper = new DispToggleButtonEvents_SinkHelper(this, _connectPoint);
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
                        Factory.Console.WriteException(exception);
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
			if( null != __ToggleButtonInOptionEvents_SinkHelper)
			{
				__ToggleButtonInOptionEvents_SinkHelper.Dispose();
				__ToggleButtonInOptionEvents_SinkHelper = null;
			}
			if( null != _dispToggleButtonEvents_SinkHelper)
			{
				_dispToggleButtonEvents_SinkHelper.Dispose();
				_dispToggleButtonEvents_SinkHelper = null;
			}

			_connectPoint = null;
		}
        
        #endregion

		#pragma warning restore
	}
}