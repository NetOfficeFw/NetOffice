using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice;
namespace NetOffice.OutlookApi
{

	#region Delegates

	#pragma warning disable
	public delegate void OlkTimeControl_ClickEventHandler();
	public delegate void OlkTimeControl_DoubleClickEventHandler();
	public delegate void OlkTimeControl_MouseDownEventHandler(NetOffice.OutlookApi.Enums.OlMouseButton Button, NetOffice.OutlookApi.Enums.OlShiftState Shift, Single X, Single Y);
	public delegate void OlkTimeControl_MouseMoveEventHandler(NetOffice.OutlookApi.Enums.OlMouseButton Button, NetOffice.OutlookApi.Enums.OlShiftState Shift, Single X, Single Y);
	public delegate void OlkTimeControl_MouseUpEventHandler(NetOffice.OutlookApi.Enums.OlMouseButton Button, NetOffice.OutlookApi.Enums.OlShiftState Shift, Single X, Single Y);
	public delegate void OlkTimeControl_EnterEventHandler();
	public delegate void OlkTimeControl_ExitEventHandler(ref bool Cancel);
	public delegate void OlkTimeControl_KeyDownEventHandler(ref Int32 KeyCode, NetOffice.OutlookApi.Enums.OlShiftState Shift);
	public delegate void OlkTimeControl_KeyPressEventHandler(ref Int32 KeyAscii);
	public delegate void OlkTimeControl_KeyUpEventHandler(ref Int32 KeyCode, NetOffice.OutlookApi.Enums.OlShiftState Shift);
	public delegate void OlkTimeControl_ChangeEventHandler();
	public delegate void OlkTimeControl_DropButtonClickEventHandler();
	public delegate void OlkTimeControl_AfterUpdateEventHandler();
	public delegate void OlkTimeControl_BeforeUpdateEventHandler(ref bool Cancel);
	#pragma warning restore

	#endregion

	///<summary>
	/// CoClass OlkTimeControl 
	/// SupportByVersion Outlook, 12,14,15
	/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff868612.aspx
	///</summary>
	[SupportByVersionAttribute("Outlook", 12,14,15)]
	[EntityTypeAttribute(EntityType.IsCoClass)]
	public class OlkTimeControl : _OlkTimeControl,IEventBinding
	{
		#pragma warning disable
		#region Fields
		
		private NetRuntimeSystem.Runtime.InteropServices.ComTypes.IConnectionPoint _connectPoint;
		private string _activeSinkId;
		private NetRuntimeSystem.Type _thisType;
		OlkTimeControlEvents_SinkHelper _olkTimeControlEvents_SinkHelper;
	
		#endregion

		#region Type Information

        private static Type _type;
		
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        public static Type LateBindingApiWrapperType
        {
            get
            {
                if (null == _type)
                    _type = typeof(OlkTimeControl);
                return _type;
            }
        }
        
        #endregion
        		
		#region Construction

		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
		public OlkTimeControl(Core factory, COMObject parentObject, object comProxy) : base(factory, parentObject, comProxy)
		{
			
		}

        ///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
		public OlkTimeControl(COMObject parentObject, object comProxy) : base(parentObject, comProxy)
		{
			
		}

		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public OlkTimeControl(Core factory, COMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(factory, parentObject, comProxy, comProxyType)
		{
			
		}

		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public OlkTimeControl(COMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(parentObject, comProxy, comProxyType)
		{
			
		}
		
		///<param name="replacedObject">object to replaced. replacedObject are not usable after this action</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public OlkTimeControl(COMObject replacedObject) : base(replacedObject)
		{
			
		}
		
		///<summary>
        ///creates a new instance of OlkTimeControl 
        ///</summary>		
		public OlkTimeControl():base("Outlook.OlkTimeControl")
		{
			
		}
		
		///<summary>
        ///creates a new instance of OlkTimeControl
        ///</summary>
        ///<param name="progId">registered ProgID</param>
		public OlkTimeControl(string progId):base(progId)
		{
			
		}

		#endregion

		#region Static CoClass Methods

		/// <summary>
        /// returns all running Outlook.OlkTimeControl objects from the running object table(ROT)
        /// </summary>
        /// <returns>an Outlook.OlkTimeControl array</returns>
		public static NetOffice.OutlookApi.OlkTimeControl[] GetActiveInstances()
		{		
			NetRuntimeSystem.Collections.Generic.List<object> proxyList = NetOffice.RunningObjectTable.GetActiveProxiesFromROT("Outlook","OlkTimeControl");
			NetRuntimeSystem.Collections.Generic.List<NetOffice.OutlookApi.OlkTimeControl> resultList = new NetRuntimeSystem.Collections.Generic.List<NetOffice.OutlookApi.OlkTimeControl>();
			foreach(object proxy in proxyList)
				resultList.Add( new NetOffice.OutlookApi.OlkTimeControl(null, proxy) );
			return resultList.ToArray();
		}

		/// <summary>
        /// returns a running Outlook.OlkTimeControl object from the running object table(ROT). the method takes the first element from the table
        /// </summary>
        /// <returns>an Outlook.OlkTimeControl object or null</returns>
		public static NetOffice.OutlookApi.OlkTimeControl GetActiveInstance()
		{
			object proxy = NetOffice.RunningObjectTable.GetActiveProxyFromROT("Outlook","OlkTimeControl", false);
			if(null != proxy)
				return new NetOffice.OutlookApi.OlkTimeControl(null, proxy);
			else
				return null;
		}

		/// <summary>
        /// returns a running Outlook.OlkTimeControl object from the running object table(ROT).  the method takes the first element from the table
        /// </summary>
	    /// <param name="throwOnError">throw an exception if no object was found</param>
        /// <returns>an Outlook.OlkTimeControl object or null</returns>
		public static NetOffice.OutlookApi.OlkTimeControl GetActiveInstance(bool throwOnError)
		{
			object proxy = NetOffice.RunningObjectTable.GetActiveProxyFromROT("Outlook","OlkTimeControl", throwOnError);
			if(null != proxy)
				return new NetOffice.OutlookApi.OlkTimeControl(null, proxy);
			else
				return null;
		}
		#endregion

		#region Events

		/// <summary>
		/// SupportByVersion Outlook, 12,14,15
		/// </summary>
		private event OlkTimeControl_ClickEventHandler _ClickEvent;

		/// <summary>
		/// SupportByVersion Outlook 12 14 15
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff866709.aspx </remarks>
		[SupportByVersion("Outlook", 12,14,15)]
		public event OlkTimeControl_ClickEventHandler ClickEvent
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
		/// SupportByVersion Outlook, 12,14,15
		/// </summary>
		private event OlkTimeControl_DoubleClickEventHandler _DoubleClickEvent;

		/// <summary>
		/// SupportByVersion Outlook 12 14 15
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff869446.aspx </remarks>
		[SupportByVersion("Outlook", 12,14,15)]
		public event OlkTimeControl_DoubleClickEventHandler DoubleClickEvent
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
		/// SupportByVersion Outlook, 12,14,15
		/// </summary>
		private event OlkTimeControl_MouseDownEventHandler _MouseDownEvent;

		/// <summary>
		/// SupportByVersion Outlook 12 14 15
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff865862.aspx </remarks>
		[SupportByVersion("Outlook", 12,14,15)]
		public event OlkTimeControl_MouseDownEventHandler MouseDownEvent
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
		/// SupportByVersion Outlook, 12,14,15
		/// </summary>
		private event OlkTimeControl_MouseMoveEventHandler _MouseMoveEvent;

		/// <summary>
		/// SupportByVersion Outlook 12 14 15
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff865094.aspx </remarks>
		[SupportByVersion("Outlook", 12,14,15)]
		public event OlkTimeControl_MouseMoveEventHandler MouseMoveEvent
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
		/// SupportByVersion Outlook, 12,14,15
		/// </summary>
		private event OlkTimeControl_MouseUpEventHandler _MouseUpEvent;

		/// <summary>
		/// SupportByVersion Outlook 12 14 15
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff870088.aspx </remarks>
		[SupportByVersion("Outlook", 12,14,15)]
		public event OlkTimeControl_MouseUpEventHandler MouseUpEvent
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
		/// SupportByVersion Outlook, 12,14,15
		/// </summary>
		private event OlkTimeControl_EnterEventHandler _EnterEvent;

		/// <summary>
		/// SupportByVersion Outlook 12 14 15
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff862681.aspx </remarks>
		[SupportByVersion("Outlook", 12,14,15)]
		public event OlkTimeControl_EnterEventHandler EnterEvent
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
		/// SupportByVersion Outlook, 12,14,15
		/// </summary>
		private event OlkTimeControl_ExitEventHandler _ExitEvent;

		/// <summary>
		/// SupportByVersion Outlook 12 14 15
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff860380.aspx </remarks>
		[SupportByVersion("Outlook", 12,14,15)]
		public event OlkTimeControl_ExitEventHandler ExitEvent
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
		/// SupportByVersion Outlook, 12,14,15
		/// </summary>
		private event OlkTimeControl_KeyDownEventHandler _KeyDownEvent;

		/// <summary>
		/// SupportByVersion Outlook 12 14 15
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff861291.aspx </remarks>
		[SupportByVersion("Outlook", 12,14,15)]
		public event OlkTimeControl_KeyDownEventHandler KeyDownEvent
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
		/// SupportByVersion Outlook, 12,14,15
		/// </summary>
		private event OlkTimeControl_KeyPressEventHandler _KeyPressEvent;

		/// <summary>
		/// SupportByVersion Outlook 12 14 15
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff865313.aspx </remarks>
		[SupportByVersion("Outlook", 12,14,15)]
		public event OlkTimeControl_KeyPressEventHandler KeyPressEvent
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
		/// SupportByVersion Outlook, 12,14,15
		/// </summary>
		private event OlkTimeControl_KeyUpEventHandler _KeyUpEvent;

		/// <summary>
		/// SupportByVersion Outlook 12 14 15
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff868629.aspx </remarks>
		[SupportByVersion("Outlook", 12,14,15)]
		public event OlkTimeControl_KeyUpEventHandler KeyUpEvent
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
		/// SupportByVersion Outlook, 12,14,15
		/// </summary>
		private event OlkTimeControl_ChangeEventHandler _ChangeEvent;

		/// <summary>
		/// SupportByVersion Outlook 12 14 15
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff867578.aspx </remarks>
		[SupportByVersion("Outlook", 12,14,15)]
		public event OlkTimeControl_ChangeEventHandler ChangeEvent
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
		/// SupportByVersion Outlook, 12,14,15
		/// </summary>
		private event OlkTimeControl_DropButtonClickEventHandler _DropButtonClickEvent;

		/// <summary>
		/// SupportByVersion Outlook 12 14 15
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff862791.aspx </remarks>
		[SupportByVersion("Outlook", 12,14,15)]
		public event OlkTimeControl_DropButtonClickEventHandler DropButtonClickEvent
		{
			add
			{
				CreateEventBridge();
				_DropButtonClickEvent += value;
			}
			remove
			{
				_DropButtonClickEvent -= value;
			}
		}

		/// <summary>
		/// SupportByVersion Outlook, 12,14,15
		/// </summary>
		private event OlkTimeControl_AfterUpdateEventHandler _AfterUpdateEvent;

		/// <summary>
		/// SupportByVersion Outlook 12 14 15
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff865068.aspx </remarks>
		[SupportByVersion("Outlook", 12,14,15)]
		public event OlkTimeControl_AfterUpdateEventHandler AfterUpdateEvent
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
		/// SupportByVersion Outlook, 12,14,15
		/// </summary>
		private event OlkTimeControl_BeforeUpdateEventHandler _BeforeUpdateEvent;

		/// <summary>
		/// SupportByVersion Outlook 12 14 15
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff868825.aspx </remarks>
		[SupportByVersion("Outlook", 12,14,15)]
		public event OlkTimeControl_BeforeUpdateEventHandler BeforeUpdateEvent
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
			if(false == Factory.Settings.EnableEvents)
				return;
	
			if (null != _connectPoint)
				return;
	
            if (null == _activeSinkId)
				_activeSinkId = SinkHelper.GetConnectionPoint(this, ref _connectPoint, OlkTimeControlEvents_SinkHelper.Id);


			if(OlkTimeControlEvents_SinkHelper.Id.Equals(_activeSinkId, StringComparison.InvariantCultureIgnoreCase))
			{
				_olkTimeControlEvents_SinkHelper = new OlkTimeControlEvents_SinkHelper(this, _connectPoint);
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
			if( null != _olkTimeControlEvents_SinkHelper)
			{
				_olkTimeControlEvents_SinkHelper.Dispose();
				_olkTimeControlEvents_SinkHelper = null;
			}

			_connectPoint = null;
		}
        
        #endregion

		#pragma warning restore
	}
}