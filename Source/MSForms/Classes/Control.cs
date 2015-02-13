using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice;
namespace NetOffice.MSFormsApi
{

	#region Delegates

	#pragma warning disable
	public delegate void Control_EnterEventHandler();
	public delegate void Control_ExitEventHandler(NetOffice.MSFormsApi.ReturnBoolean Cancel);
	public delegate void Control_BeforeUpdateEventHandler(NetOffice.MSFormsApi.ReturnBoolean Cancel);
	public delegate void Control_AfterUpdateEventHandler();
	#pragma warning restore

	#endregion

	///<summary>
	/// CoClass Control 
	/// SupportByVersion MSForms, 2
	///</summary>
	[SupportByVersionAttribute("MSForms", 2)]
	[EntityTypeAttribute(EntityType.IsCoClass)]
	public class Control : IControl,IEventBinding
	{
		#pragma warning disable
		#region Fields
		
		private NetRuntimeSystem.Runtime.InteropServices.ComTypes.IConnectionPoint _connectPoint;
		private string _activeSinkId;
		private NetRuntimeSystem.Type _thisType;
		ControlEvents_SinkHelper _controlEvents_SinkHelper;
	
		#endregion

		#region Type Information

        private static Type _type;
		
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        public static Type LateBindingApiWrapperType
        {
            get
            {
                if (null == _type)
                    _type = typeof(Control);
                return _type;
            }
        }
        
        #endregion
        		
		#region Construction

		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
		public Control(Core factory, COMObject parentObject, object comProxy) : base(factory, parentObject, comProxy)
		{
			
		}

        ///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
		public Control(COMObject parentObject, object comProxy) : base(parentObject, comProxy)
		{
			
		}

		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Control(Core factory, COMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(factory, parentObject, comProxy, comProxyType)
		{
			
		}

		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Control(COMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(parentObject, comProxy, comProxyType)
		{
			
		}
		
		///<param name="replacedObject">object to replaced. replacedObject are not usable after this action</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Control(COMObject replacedObject) : base(replacedObject)
		{
			
		}
		
		///<summary>
        ///creates a new instance of Control 
        ///</summary>		
		public Control():base("MSForms.Control")
		{
			
		}
		
		///<summary>
        ///creates a new instance of Control
        ///</summary>
        ///<param name="progId">registered ProgID</param>
		public Control(string progId):base(progId)
		{
			
		}

		#endregion

		#region Static CoClass Methods

		/// <summary>
        /// returns all running MSForms.Control objects from the running object table(ROT)
        /// </summary>
        /// <returns>an MSForms.Control array</returns>
		public static NetOffice.MSFormsApi.Control[] GetActiveInstances()
		{		
			NetRuntimeSystem.Collections.Generic.List<object> proxyList = NetOffice.RunningObjectTable.GetActiveProxiesFromROT("MSForms","Control");
			NetRuntimeSystem.Collections.Generic.List<NetOffice.MSFormsApi.Control> resultList = new NetRuntimeSystem.Collections.Generic.List<NetOffice.MSFormsApi.Control>();
			foreach(object proxy in proxyList)
				resultList.Add( new NetOffice.MSFormsApi.Control(null, proxy) );
			return resultList.ToArray();
		}

		/// <summary>
        /// returns a running MSForms.Control object from the running object table(ROT). the method takes the first element from the table
        /// </summary>
        /// <returns>an MSForms.Control object or null</returns>
		public static NetOffice.MSFormsApi.Control GetActiveInstance()
		{
			object proxy = NetOffice.RunningObjectTable.GetActiveProxyFromROT("MSForms","Control", false);
			if(null != proxy)
				return new NetOffice.MSFormsApi.Control(null, proxy);
			else
				return null;
		}

		/// <summary>
        /// returns a running MSForms.Control object from the running object table(ROT).  the method takes the first element from the table
        /// </summary>
	    /// <param name="throwOnError">throw an exception if no object was found</param>
        /// <returns>an MSForms.Control object or null</returns>
		public static NetOffice.MSFormsApi.Control GetActiveInstance(bool throwOnError)
		{
			object proxy = NetOffice.RunningObjectTable.GetActiveProxyFromROT("MSForms","Control", throwOnError);
			if(null != proxy)
				return new NetOffice.MSFormsApi.Control(null, proxy);
			else
				return null;
		}
		#endregion

		#region Events

		/// <summary>
		/// SupportByVersion MSForms, 2
		/// </summary>
		private event Control_EnterEventHandler _EnterEvent;

		/// <summary>
		/// SupportByVersion MSForms 2
		/// </summary>
		[SupportByVersion("MSForms", 2)]
		public event Control_EnterEventHandler EnterEvent
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
		/// SupportByVersion MSForms, 2
		/// </summary>
		private event Control_ExitEventHandler _ExitEvent;

		/// <summary>
		/// SupportByVersion MSForms 2
		/// </summary>
		[SupportByVersion("MSForms", 2)]
		public event Control_ExitEventHandler ExitEvent
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
		/// SupportByVersion MSForms, 2
		/// </summary>
		private event Control_BeforeUpdateEventHandler _BeforeUpdateEvent;

		/// <summary>
		/// SupportByVersion MSForms 2
		/// </summary>
		[SupportByVersion("MSForms", 2)]
		public event Control_BeforeUpdateEventHandler BeforeUpdateEvent
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
		/// SupportByVersion MSForms, 2
		/// </summary>
		private event Control_AfterUpdateEventHandler _AfterUpdateEvent;

		/// <summary>
		/// SupportByVersion MSForms 2
		/// </summary>
		[SupportByVersion("MSForms", 2)]
		public event Control_AfterUpdateEventHandler AfterUpdateEvent
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
				_activeSinkId = SinkHelper.GetConnectionPoint(this, ref _connectPoint, ControlEvents_SinkHelper.Id);


			if(ControlEvents_SinkHelper.Id.Equals(_activeSinkId, StringComparison.InvariantCultureIgnoreCase))
			{
				_controlEvents_SinkHelper = new ControlEvents_SinkHelper(this, _connectPoint);
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
			if( null != _controlEvents_SinkHelper)
			{
				_controlEvents_SinkHelper.Dispose();
				_controlEvents_SinkHelper = null;
			}

			_connectPoint = null;
		}
        
        #endregion

		#pragma warning restore
	}
}