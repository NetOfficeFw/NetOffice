using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice;
namespace NetOffice.VisioApi
{

	#region Delegates

	#pragma warning disable
	public delegate void Style_StyleChangedEventHandler(NetOffice.VisioApi.IVStyle Style);
	public delegate void Style_BeforeStyleDeleteEventHandler(NetOffice.VisioApi.IVStyle Style);
	public delegate void Style_QueryCancelStyleDeleteEventHandler(NetOffice.VisioApi.IVStyle Style);
	public delegate void Style_StyleDeleteCanceledEventHandler(NetOffice.VisioApi.IVStyle Style);
	#pragma warning restore

	#endregion

	///<summary>
	/// CoClass Style 
	/// SupportByVersion Visio, 11,12,14,15
	/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/ff769398(v=office.14).aspx
	///</summary>
	[SupportByVersionAttribute("Visio", 11,12,14,15)]
	[EntityTypeAttribute(EntityType.IsCoClass)]
	public class Style : IVStyle,IEventBinding
	{
		#pragma warning disable
		#region Fields
		
		private NetRuntimeSystem.Runtime.InteropServices.ComTypes.IConnectionPoint _connectPoint;
		private string _activeSinkId;
		private NetRuntimeSystem.Type _thisType;
		EStyle_SinkHelper _eStyle_SinkHelper;
	
		#endregion

		#region Type Information

        private static Type _type;
		
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        public static Type LateBindingApiWrapperType
        {
            get
            {
                if (null == _type)
                    _type = typeof(Style);
                return _type;
            }
        }
        
        #endregion
        		
		#region Construction

		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
		public Style(Core factory, COMObject parentObject, object comProxy) : base(factory, parentObject, comProxy)
		{
			
		}

        ///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
		public Style(COMObject parentObject, object comProxy) : base(parentObject, comProxy)
		{
			
		}

		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Style(Core factory, COMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(factory, parentObject, comProxy, comProxyType)
		{
			
		}

		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Style(COMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(parentObject, comProxy, comProxyType)
		{
			
		}
		
		///<param name="replacedObject">object to replaced. replacedObject are not usable after this action</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Style(COMObject replacedObject) : base(replacedObject)
		{
			
		}
		
		///<summary>
        ///creates a new instance of Style 
        ///</summary>		
		public Style():base("Visio.Style")
		{
			
		}
		
		///<summary>
        ///creates a new instance of Style
        ///</summary>
        ///<param name="progId">registered ProgID</param>
		public Style(string progId):base(progId)
		{
			
		}

		#endregion

		#region Static CoClass Methods

		/// <summary>
        /// returns all running Visio.Style objects from the running object table(ROT)
        /// </summary>
        /// <returns>an Visio.Style array</returns>
		public static NetOffice.VisioApi.Style[] GetActiveInstances()
		{		
			NetRuntimeSystem.Collections.Generic.List<object> proxyList = NetOffice.RunningObjectTable.GetActiveProxiesFromROT("Visio","Style");
			NetRuntimeSystem.Collections.Generic.List<NetOffice.VisioApi.Style> resultList = new NetRuntimeSystem.Collections.Generic.List<NetOffice.VisioApi.Style>();
			foreach(object proxy in proxyList)
				resultList.Add( new NetOffice.VisioApi.Style(null, proxy) );
			return resultList.ToArray();
		}

		/// <summary>
        /// returns a running Visio.Style object from the running object table(ROT). the method takes the first element from the table
        /// </summary>
        /// <returns>an Visio.Style object or null</returns>
		public static NetOffice.VisioApi.Style GetActiveInstance()
		{
			object proxy = NetOffice.RunningObjectTable.GetActiveProxyFromROT("Visio","Style", false);
			if(null != proxy)
				return new NetOffice.VisioApi.Style(null, proxy);
			else
				return null;
		}

		/// <summary>
        /// returns a running Visio.Style object from the running object table(ROT).  the method takes the first element from the table
        /// </summary>
	    /// <param name="throwOnError">throw an exception if no object was found</param>
        /// <returns>an Visio.Style object or null</returns>
		public static NetOffice.VisioApi.Style GetActiveInstance(bool throwOnError)
		{
			object proxy = NetOffice.RunningObjectTable.GetActiveProxyFromROT("Visio","Style", throwOnError);
			if(null != proxy)
				return new NetOffice.VisioApi.Style(null, proxy);
			else
				return null;
		}
		#endregion

		#region Events

		/// <summary>
		/// SupportByVersion Visio, 11,12,14,15
		/// </summary>
		private event Style_StyleChangedEventHandler _StyleChangedEvent;

		/// <summary>
		/// SupportByVersion Visio 11 12 14 15
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/ff768814(v=office.14).aspx </remarks>
		[SupportByVersion("Visio", 11,12,14,15)]
		public event Style_StyleChangedEventHandler StyleChangedEvent
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
		private event Style_BeforeStyleDeleteEventHandler _BeforeStyleDeleteEvent;

		/// <summary>
		/// SupportByVersion Visio 11 12 14 15
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/ff765885(v=office.14).aspx </remarks>
		[SupportByVersion("Visio", 11,12,14,15)]
		public event Style_BeforeStyleDeleteEventHandler BeforeStyleDeleteEvent
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
		private event Style_QueryCancelStyleDeleteEventHandler _QueryCancelStyleDeleteEvent;

		/// <summary>
		/// SupportByVersion Visio 11 12 14 15
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/ff766049(v=office.14).aspx </remarks>
		[SupportByVersion("Visio", 11,12,14,15)]
		public event Style_QueryCancelStyleDeleteEventHandler QueryCancelStyleDeleteEvent
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
		private event Style_StyleDeleteCanceledEventHandler _StyleDeleteCanceledEvent;

		/// <summary>
		/// SupportByVersion Visio 11 12 14 15
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/ff765208(v=office.14).aspx </remarks>
		[SupportByVersion("Visio", 11,12,14,15)]
		public event Style_StyleDeleteCanceledEventHandler StyleDeleteCanceledEvent
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
				_activeSinkId = SinkHelper.GetConnectionPoint(this, ref _connectPoint, EStyle_SinkHelper.Id);


			if(EStyle_SinkHelper.Id.Equals(_activeSinkId, StringComparison.InvariantCultureIgnoreCase))
			{
				_eStyle_SinkHelper = new EStyle_SinkHelper(this, _connectPoint);
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
			if( null != _eStyle_SinkHelper)
			{
				_eStyle_SinkHelper.Dispose();
				_eStyle_SinkHelper = null;
			}

			_connectPoint = null;
		}
        
        #endregion

		#pragma warning restore
	}
}