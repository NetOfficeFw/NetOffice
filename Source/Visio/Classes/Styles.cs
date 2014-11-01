using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice;
namespace NetOffice.VisioApi
{

	#region Delegates

	#pragma warning disable
	public delegate void Styles_StyleAddedEventHandler(NetOffice.VisioApi.IVStyle Style);
	public delegate void Styles_StyleChangedEventHandler(NetOffice.VisioApi.IVStyle Style);
	public delegate void Styles_BeforeStyleDeleteEventHandler(NetOffice.VisioApi.IVStyle Style);
	public delegate void Styles_QueryCancelStyleDeleteEventHandler(NetOffice.VisioApi.IVStyle Style);
	public delegate void Styles_StyleDeleteCanceledEventHandler(NetOffice.VisioApi.IVStyle Style);
	#pragma warning restore

	#endregion

	///<summary>
	/// CoClass Styles 
	/// SupportByVersion Visio, 11,12,14,15
	/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/ff769402(v=office.14).aspx
	///</summary>
	[SupportByVersionAttribute("Visio", 11,12,14,15)]
	[EntityTypeAttribute(EntityType.IsCoClass)]
	public class Styles : IVStyles,IEventBinding
	{
		#pragma warning disable
		#region Fields
		
		private NetRuntimeSystem.Runtime.InteropServices.ComTypes.IConnectionPoint _connectPoint;
		private string _activeSinkId;
		private NetRuntimeSystem.Type _thisType;
		EStyles_SinkHelper _eStyles_SinkHelper;
	
		#endregion

		#region Type Information

        private static Type _type;
		
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        public static Type LateBindingApiWrapperType
        {
            get
            {
                if (null == _type)
                    _type = typeof(Styles);
                return _type;
            }
        }
        
        #endregion
        		
		#region Construction

		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
		public Styles(Core factory, COMObject parentObject, object comProxy) : base(factory, parentObject, comProxy)
		{
			
		}

        ///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
		public Styles(COMObject parentObject, object comProxy) : base(parentObject, comProxy)
		{
			
		}

		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Styles(Core factory, COMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(factory, parentObject, comProxy, comProxyType)
		{
			
		}

		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Styles(COMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(parentObject, comProxy, comProxyType)
		{
			
		}
		
		///<param name="replacedObject">object to replaced. replacedObject are not usable after this action</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Styles(COMObject replacedObject) : base(replacedObject)
		{
			
		}
		
		///<summary>
        ///creates a new instance of Styles 
        ///</summary>		
		public Styles():base("Visio.Styles")
		{
			
		}
		
		///<summary>
        ///creates a new instance of Styles
        ///</summary>
        ///<param name="progId">registered ProgID</param>
		public Styles(string progId):base(progId)
		{
			
		}

		#endregion

		#region Static CoClass Methods

		/// <summary>
        /// returns all running Visio.Styles objects from the running object table(ROT)
        /// </summary>
        /// <returns>an Visio.Styles array</returns>
		public static NetOffice.VisioApi.Styles[] GetActiveInstances()
		{		
			NetRuntimeSystem.Collections.Generic.List<object> proxyList = NetOffice.RunningObjectTable.GetActiveProxiesFromROT("Visio","Styles");
			NetRuntimeSystem.Collections.Generic.List<NetOffice.VisioApi.Styles> resultList = new NetRuntimeSystem.Collections.Generic.List<NetOffice.VisioApi.Styles>();
			foreach(object proxy in proxyList)
				resultList.Add( new NetOffice.VisioApi.Styles(null, proxy) );
			return resultList.ToArray();
		}

		/// <summary>
        /// returns a running Visio.Styles object from the running object table(ROT). the method takes the first element from the table
        /// </summary>
        /// <returns>an Visio.Styles object or null</returns>
		public static NetOffice.VisioApi.Styles GetActiveInstance()
		{
			object proxy = NetOffice.RunningObjectTable.GetActiveProxyFromROT("Visio","Styles", false);
			if(null != proxy)
				return new NetOffice.VisioApi.Styles(null, proxy);
			else
				return null;
		}

		/// <summary>
        /// returns a running Visio.Styles object from the running object table(ROT).  the method takes the first element from the table
        /// </summary>
	    /// <param name="throwOnError">throw an exception if no object was found</param>
        /// <returns>an Visio.Styles object or null</returns>
		public static NetOffice.VisioApi.Styles GetActiveInstance(bool throwOnError)
		{
			object proxy = NetOffice.RunningObjectTable.GetActiveProxyFromROT("Visio","Styles", throwOnError);
			if(null != proxy)
				return new NetOffice.VisioApi.Styles(null, proxy);
			else
				return null;
		}
		#endregion

		#region Events

		/// <summary>
		/// SupportByVersion Visio, 11,12,14,15
		/// </summary>
		private event Styles_StyleAddedEventHandler _StyleAddedEvent;

		/// <summary>
		/// SupportByVersion Visio 11 12 14 15
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/ff768234(v=office.14).aspx </remarks>
		[SupportByVersion("Visio", 11,12,14,15)]
		public event Styles_StyleAddedEventHandler StyleAddedEvent
		{
			add
			{
				CreateEventBridge();
				_StyleAddedEvent += value;
			}
			remove
			{
				_StyleAddedEvent -= value;
			}
		}

		/// <summary>
		/// SupportByVersion Visio, 11,12,14,15
		/// </summary>
		private event Styles_StyleChangedEventHandler _StyleChangedEvent;

		/// <summary>
		/// SupportByVersion Visio 11 12 14 15
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/ff766489(v=office.14).aspx </remarks>
		[SupportByVersion("Visio", 11,12,14,15)]
		public event Styles_StyleChangedEventHandler StyleChangedEvent
		{
			add
			{
				CreateEventBridge();
				_StyleChangedEvent += value;
			}
			remove
			{
				_StyleChangedEvent -= value;
			}
		}

		/// <summary>
		/// SupportByVersion Visio, 11,12,14,15
		/// </summary>
		private event Styles_BeforeStyleDeleteEventHandler _BeforeStyleDeleteEvent;

		/// <summary>
		/// SupportByVersion Visio 11 12 14 15
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/ff768770(v=office.14).aspx </remarks>
		[SupportByVersion("Visio", 11,12,14,15)]
		public event Styles_BeforeStyleDeleteEventHandler BeforeStyleDeleteEvent
		{
			add
			{
				CreateEventBridge();
				_BeforeStyleDeleteEvent += value;
			}
			remove
			{
				_BeforeStyleDeleteEvent -= value;
			}
		}

		/// <summary>
		/// SupportByVersion Visio, 11,12,14,15
		/// </summary>
		private event Styles_QueryCancelStyleDeleteEventHandler _QueryCancelStyleDeleteEvent;

		/// <summary>
		/// SupportByVersion Visio 11 12 14 15
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/ff765172(v=office.14).aspx </remarks>
		[SupportByVersion("Visio", 11,12,14,15)]
		public event Styles_QueryCancelStyleDeleteEventHandler QueryCancelStyleDeleteEvent
		{
			add
			{
				CreateEventBridge();
				_QueryCancelStyleDeleteEvent += value;
			}
			remove
			{
				_QueryCancelStyleDeleteEvent -= value;
			}
		}

		/// <summary>
		/// SupportByVersion Visio, 11,12,14,15
		/// </summary>
		private event Styles_StyleDeleteCanceledEventHandler _StyleDeleteCanceledEvent;

		/// <summary>
		/// SupportByVersion Visio 11 12 14 15
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/ff767143(v=office.14).aspx </remarks>
		[SupportByVersion("Visio", 11,12,14,15)]
		public event Styles_StyleDeleteCanceledEventHandler StyleDeleteCanceledEvent
		{
			add
			{
				CreateEventBridge();
				_StyleDeleteCanceledEvent += value;
			}
			remove
			{
				_StyleDeleteCanceledEvent -= value;
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
				_activeSinkId = SinkHelper.GetConnectionPoint(this, ref _connectPoint, EStyles_SinkHelper.Id);


			if(EStyles_SinkHelper.Id.Equals(_activeSinkId, StringComparison.InvariantCultureIgnoreCase))
			{
				_eStyles_SinkHelper = new EStyles_SinkHelper(this, _connectPoint);
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
			if( null != _eStyles_SinkHelper)
			{
				_eStyles_SinkHelper.Dispose();
				_eStyles_SinkHelper = null;
			}

			_connectPoint = null;
		}
        
        #endregion

		#pragma warning restore
	}
}