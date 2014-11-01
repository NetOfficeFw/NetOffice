using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice;
namespace NetOffice.AccessApi
{

	#region Delegates

	#pragma warning disable
	public delegate void CommandButton_ClickEventHandler();
	public delegate void CommandButton_EnterEventHandler();
	public delegate void CommandButton_ExitEventHandler(ref Int16 Cancel);
	public delegate void CommandButton_GotFocusEventHandler();
	public delegate void CommandButton_LostFocusEventHandler();
	public delegate void CommandButton_DblClickEventHandler(ref Int16 Cancel);
	public delegate void CommandButton_MouseDownEventHandler(ref Int16 Button, ref Int16 Shift, ref Single X, ref Single Y);
	public delegate void CommandButton_MouseMoveEventHandler(ref Int16 Button, ref Int16 Shift, ref Single X, ref Single Y);
	public delegate void CommandButton_MouseUpEventHandler(ref Int16 Button, ref Int16 Shift, ref Single X, ref Single Y);
	public delegate void CommandButton_KeyDownEventHandler(ref Int16 KeyCode, ref Int16 Shift);
	public delegate void CommandButton_KeyPressEventHandler(ref Int16 KeyAscii);
	public delegate void CommandButton_KeyUpEventHandler(ref Int16 KeyCode, ref Int16 Shift);
	#pragma warning restore

	#endregion

	///<summary>
	/// CoClass CommandButton 
	/// SupportByVersion Access, 9,10,11,12,14,15
	/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff191876.aspx
	///</summary>
	[SupportByVersionAttribute("Access", 9,10,11,12,14,15)]
	[EntityTypeAttribute(EntityType.IsCoClass)]
	public class CommandButton : _CommandButton,IEventBinding
	{
		#pragma warning disable
		#region Fields
		
		private NetRuntimeSystem.Runtime.InteropServices.ComTypes.IConnectionPoint _connectPoint;
		private string _activeSinkId;
		private NetRuntimeSystem.Type _thisType;
		_CommandButtonEvents_SinkHelper __CommandButtonEvents_SinkHelper;
		DispCommandButtonEvents_SinkHelper _dispCommandButtonEvents_SinkHelper;
	
		#endregion

		#region Type Information

        private static Type _type;
		
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        public static Type LateBindingApiWrapperType
        {
            get
            {
                if (null == _type)
                    _type = typeof(CommandButton);
                return _type;
            }
        }
        
        #endregion
        		
		#region Construction

		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
		public CommandButton(Core factory, COMObject parentObject, object comProxy) : base(factory, parentObject, comProxy)
		{
			
		}

        ///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
		public CommandButton(COMObject parentObject, object comProxy) : base(parentObject, comProxy)
		{
			
		}

		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public CommandButton(Core factory, COMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(factory, parentObject, comProxy, comProxyType)
		{
			
		}

		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public CommandButton(COMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(parentObject, comProxy, comProxyType)
		{
			
		}
		
		///<param name="replacedObject">object to replaced. replacedObject are not usable after this action</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public CommandButton(COMObject replacedObject) : base(replacedObject)
		{
			
		}
		
		///<summary>
        ///creates a new instance of CommandButton 
        ///</summary>		
		public CommandButton():base("Access.CommandButton")
		{
			
		}
		
		///<summary>
        ///creates a new instance of CommandButton
        ///</summary>
        ///<param name="progId">registered ProgID</param>
		public CommandButton(string progId):base(progId)
		{
			
		}

		#endregion

		#region Static CoClass Methods

		/// <summary>
        /// returns all running Access.CommandButton objects from the running object table(ROT)
        /// </summary>
        /// <returns>an Access.CommandButton array</returns>
		public static NetOffice.AccessApi.CommandButton[] GetActiveInstances()
		{		
			NetRuntimeSystem.Collections.Generic.List<object> proxyList = NetOffice.RunningObjectTable.GetActiveProxiesFromROT("Access","CommandButton");
			NetRuntimeSystem.Collections.Generic.List<NetOffice.AccessApi.CommandButton> resultList = new NetRuntimeSystem.Collections.Generic.List<NetOffice.AccessApi.CommandButton>();
			foreach(object proxy in proxyList)
				resultList.Add( new NetOffice.AccessApi.CommandButton(null, proxy) );
			return resultList.ToArray();
		}

		/// <summary>
        /// returns a running Access.CommandButton object from the running object table(ROT). the method takes the first element from the table
        /// </summary>
        /// <returns>an Access.CommandButton object or null</returns>
		public static NetOffice.AccessApi.CommandButton GetActiveInstance()
		{
			object proxy = NetOffice.RunningObjectTable.GetActiveProxyFromROT("Access","CommandButton", false);
			if(null != proxy)
				return new NetOffice.AccessApi.CommandButton(null, proxy);
			else
				return null;
		}

		/// <summary>
        /// returns a running Access.CommandButton object from the running object table(ROT).  the method takes the first element from the table
        /// </summary>
	    /// <param name="throwOnError">throw an exception if no object was found</param>
        /// <returns>an Access.CommandButton object or null</returns>
		public static NetOffice.AccessApi.CommandButton GetActiveInstance(bool throwOnError)
		{
			object proxy = NetOffice.RunningObjectTable.GetActiveProxyFromROT("Access","CommandButton", throwOnError);
			if(null != proxy)
				return new NetOffice.AccessApi.CommandButton(null, proxy);
			else
				return null;
		}
		#endregion

		#region Events

		/// <summary>
		/// SupportByVersion Access, 9,10,11,12,14,15
		/// </summary>
		private event CommandButton_ClickEventHandler _ClickEvent;

		/// <summary>
		/// SupportByVersion Access 9 10 11 12 14 15
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff822439.aspx </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15)]
		public event CommandButton_ClickEventHandler ClickEvent
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
		/// SupportByVersion Access, 9,10,11,12,14,15
		/// </summary>
		private event CommandButton_EnterEventHandler _EnterEvent;

		/// <summary>
		/// SupportByVersion Access 9 10 11 12 14 15
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff834424.aspx </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15)]
		public event CommandButton_EnterEventHandler EnterEvent
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
		/// SupportByVersion Access, 9,10,11,12,14,15
		/// </summary>
		private event CommandButton_ExitEventHandler _ExitEvent;

		/// <summary>
		/// SupportByVersion Access 9 10 11 12 14 15
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff834787.aspx </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15)]
		public event CommandButton_ExitEventHandler ExitEvent
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
		/// SupportByVersion Access, 9,10,11,12,14,15
		/// </summary>
		private event CommandButton_GotFocusEventHandler _GotFocusEvent;

		/// <summary>
		/// SupportByVersion Access 9 10 11 12 14 15
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff822451.aspx </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15)]
		public event CommandButton_GotFocusEventHandler GotFocusEvent
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
		private event CommandButton_LostFocusEventHandler _LostFocusEvent;

		/// <summary>
		/// SupportByVersion Access 9 10 11 12 14 15
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff821407.aspx </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15)]
		public event CommandButton_LostFocusEventHandler LostFocusEvent
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
		private event CommandButton_DblClickEventHandler _DblClickEvent;

		/// <summary>
		/// SupportByVersion Access 9 10 11 12 14 15
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff845152.aspx </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15)]
		public event CommandButton_DblClickEventHandler DblClickEvent
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
		/// SupportByVersion Access, 9,10,11,12,14,15
		/// </summary>
		private event CommandButton_MouseDownEventHandler _MouseDownEvent;

		/// <summary>
		/// SupportByVersion Access 9 10 11 12 14 15
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff197358.aspx </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15)]
		public event CommandButton_MouseDownEventHandler MouseDownEvent
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
		private event CommandButton_MouseMoveEventHandler _MouseMoveEvent;

		/// <summary>
		/// SupportByVersion Access 9 10 11 12 14 15
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff836610.aspx </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15)]
		public event CommandButton_MouseMoveEventHandler MouseMoveEvent
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
		private event CommandButton_MouseUpEventHandler _MouseUpEvent;

		/// <summary>
		/// SupportByVersion Access 9 10 11 12 14 15
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff197657.aspx </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15)]
		public event CommandButton_MouseUpEventHandler MouseUpEvent
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
		private event CommandButton_KeyDownEventHandler _KeyDownEvent;

		/// <summary>
		/// SupportByVersion Access 9 10 11 12 14 15
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff834776.aspx </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15)]
		public event CommandButton_KeyDownEventHandler KeyDownEvent
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
		private event CommandButton_KeyPressEventHandler _KeyPressEvent;

		/// <summary>
		/// SupportByVersion Access 9 10 11 12 14 15
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff821769.aspx </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15)]
		public event CommandButton_KeyPressEventHandler KeyPressEvent
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
		private event CommandButton_KeyUpEventHandler _KeyUpEvent;

		/// <summary>
		/// SupportByVersion Access 9 10 11 12 14 15
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff195077.aspx </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15)]
		public event CommandButton_KeyUpEventHandler KeyUpEvent
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
				_activeSinkId = SinkHelper.GetConnectionPoint(this, ref _connectPoint, _CommandButtonEvents_SinkHelper.Id,DispCommandButtonEvents_SinkHelper.Id);


			if(_CommandButtonEvents_SinkHelper.Id.Equals(_activeSinkId, StringComparison.InvariantCultureIgnoreCase))
			{
				__CommandButtonEvents_SinkHelper = new _CommandButtonEvents_SinkHelper(this, _connectPoint);
				return;
			}

			if(DispCommandButtonEvents_SinkHelper.Id.Equals(_activeSinkId, StringComparison.InvariantCultureIgnoreCase))
			{
				_dispCommandButtonEvents_SinkHelper = new DispCommandButtonEvents_SinkHelper(this, _connectPoint);
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
			if( null != __CommandButtonEvents_SinkHelper)
			{
				__CommandButtonEvents_SinkHelper.Dispose();
				__CommandButtonEvents_SinkHelper = null;
			}
			if( null != _dispCommandButtonEvents_SinkHelper)
			{
				_dispCommandButtonEvents_SinkHelper.Dispose();
				_dispCommandButtonEvents_SinkHelper = null;
			}

			_connectPoint = null;
		}
        
        #endregion

		#pragma warning restore
	}
}