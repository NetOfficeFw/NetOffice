using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice;
using NetOffice.Misc;

namespace NetOffice.ADODBApi
{

	#region Delegates

	#pragma warning disable
	public delegate void Connection_InfoMessageEventHandler(NetOffice.ADODBApi.Error pError, NetOffice.ADODBApi.Enums.EventStatusEnum adStatus, NetOffice.ADODBApi._Connection pConnection);
	public delegate void Connection_BeginTransCompleteEventHandler(Int32 TransactionLevel, NetOffice.ADODBApi.Error pError, NetOffice.ADODBApi.Enums.EventStatusEnum adStatus, NetOffice.ADODBApi._Connection pConnection);
	public delegate void Connection_CommitTransCompleteEventHandler(NetOffice.ADODBApi.Error pError, NetOffice.ADODBApi.Enums.EventStatusEnum adStatus, NetOffice.ADODBApi._Connection pConnection);
	public delegate void Connection_RollbackTransCompleteEventHandler(NetOffice.ADODBApi.Error pError, NetOffice.ADODBApi.Enums.EventStatusEnum adStatus, NetOffice.ADODBApi._Connection pConnection);
	public delegate void Connection_WillExecuteEventHandler(ref string Source, NetOffice.ADODBApi.Enums.CursorTypeEnum CursorType, NetOffice.ADODBApi.Enums.LockTypeEnum LockType, ref Int32 Options, NetOffice.ADODBApi.Enums.EventStatusEnum adStatus, NetOffice.ADODBApi._Command pCommand, NetOffice.ADODBApi._Recordset pRecordset, NetOffice.ADODBApi._Connection pConnection);
	public delegate void Connection_ExecuteCompleteEventHandler(Int32 RecordsAffected, NetOffice.ADODBApi.Error pError, NetOffice.ADODBApi.Enums.EventStatusEnum adStatus, NetOffice.ADODBApi._Command pCommand, NetOffice.ADODBApi._Recordset pRecordset, NetOffice.ADODBApi._Connection pConnection);
	public delegate void Connection_WillConnectEventHandler(ref string ConnectionString, ref string UserID, ref string Password, ref Int32 Options, NetOffice.ADODBApi.Enums.EventStatusEnum adStatus, NetOffice.ADODBApi._Connection pConnection);
	public delegate void Connection_ConnectCompleteEventHandler(NetOffice.ADODBApi.Error pError, NetOffice.ADODBApi.Enums.EventStatusEnum adStatus, NetOffice.ADODBApi._Connection pConnection);
	public delegate void Connection_DisconnectEventHandler(NetOffice.ADODBApi.Enums.EventStatusEnum adStatus, NetOffice.ADODBApi._Connection pConnection);
	#pragma warning restore

	#endregion

	///<summary>
	/// CoClass Connection 
	/// SupportByVersion ADODB, 2.1,2.5
	///</summary>
	[SupportByVersionAttribute("ADODB", 2.1,2.5)]
	[EntityTypeAttribute(EntityType.IsCoClass)]
	public class Connection : _Connection,IEventBinding
	{
		#pragma warning disable
		#region Fields
		
		private NetRuntimeSystem.Runtime.InteropServices.ComTypes.IConnectionPoint _connectPoint;
		private string _activeSinkId;
		private NetRuntimeSystem.Type _thisType;
		ConnectionEvents_SinkHelper _connectionEvents_SinkHelper;
	
		#endregion

		#region Type Information

        /// <summary>
        /// Instance Type
        /// </summary>
        public override Type InstanceType
        {
            get
            {
                return LateBindingApiWrapperType;
            }
        }

        private static Type _type;
		
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        public static Type LateBindingApiWrapperType
        {
            get
            {
                if (null == _type)
                    _type = typeof(Connection);
                return _type;
            }
        }
        
        #endregion
        		
		#region Construction

		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
		public Connection(Core factory, ICOMObject parentObject, object comProxy) : base(factory, parentObject, comProxy)
		{
			
		}

        ///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
		public Connection(ICOMObject parentObject, object comProxy) : base(parentObject, comProxy)
		{
			
		}

		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Connection(Core factory, ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(factory, parentObject, comProxy, comProxyType)
		{
			
		}

		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Connection(ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(parentObject, comProxy, comProxyType)
		{
			
		}
		
		///<param name="replacedObject">object to replaced. replacedObject are not usable after this action</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Connection(ICOMObject replacedObject) : base(replacedObject)
		{
			
		}
		
		///<summary>
        /// Creates a new instance of Connection 
        ///</summary>		
		public Connection():base("ADODB.Connection")
		{
			
		}
		
		///<summary>
        /// Creates a new instance of Connection
        ///</summary>
        ///<param name="progId">registered ProgID</param>
		public Connection(string progId):base(progId)
		{
			
		}

		#endregion

		#region Static CoClass Methods

		/// <summary>
        /// Returns all running ADODB.Connection objects from the environment/system
        /// </summary>
        /// <returns>an ADODB.Connection array</returns>
		public static NetOffice.ADODBApi.Connection[] GetActiveInstances()
		{		
			IDisposableEnumeration proxyList = NetOffice.ProxyService.GetActiveInstances("ADODB","Connection");
			NetRuntimeSystem.Collections.Generic.List<NetOffice.ADODBApi.Connection> resultList = new NetRuntimeSystem.Collections.Generic.List<NetOffice.ADODBApi.Connection>();
			foreach(object proxy in proxyList)
				resultList.Add( new NetOffice.ADODBApi.Connection(null, proxy) );
			return resultList.ToArray();
		}

		/// <summary>
        /// Returns a running ADODB.Connection object from the environment/system.
        /// </summary>
        /// <returns>an ADODB.Connection object or null</returns>
		public static NetOffice.ADODBApi.Connection GetActiveInstance()
		{
			object proxy  = NetOffice.ProxyService.GetActiveInstance("ADODB","Connection", false);
			if(null != proxy)
				return new NetOffice.ADODBApi.Connection(null, proxy);
			else
				return null;
		}

		/// <summary>
        /// Returns a running ADODB.Connection object from the environment/system. 
        /// </summary>
	    /// <param name="throwOnError">throw an exception if no object was found</param>
        /// <returns>an ADODB.Connection object or null</returns>
		public static NetOffice.ADODBApi.Connection GetActiveInstance(bool throwOnError)
		{
			object proxy  = NetOffice.ProxyService.GetActiveInstance("ADODB","Connection", throwOnError);
			if(null != proxy)
				return new NetOffice.ADODBApi.Connection(null, proxy);
			else
				return null;
		}
		#endregion

		#region Events

		/// <summary>
		/// SupportByVersion ADODB, 2.1,2.5
		/// </summary>
		private event Connection_InfoMessageEventHandler _InfoMessageEvent;

		/// <summary>
		/// SupportByVersion ADODB 2.1 2.5
		/// </summary>
		[SupportByVersion("ADODB", 2.1,2.5)]
		public event Connection_InfoMessageEventHandler InfoMessageEvent
		{
			add
			{
				CreateEventBridge();
				_InfoMessageEvent += value;
			}
			remove
			{
				_InfoMessageEvent -= value;
			}
		}

		/// <summary>
		/// SupportByVersion ADODB, 2.1,2.5
		/// </summary>
		private event Connection_BeginTransCompleteEventHandler _BeginTransCompleteEvent;

		/// <summary>
		/// SupportByVersion ADODB 2.1 2.5
		/// </summary>
		[SupportByVersion("ADODB", 2.1,2.5)]
		public event Connection_BeginTransCompleteEventHandler BeginTransCompleteEvent
		{
			add
			{
				CreateEventBridge();
				_BeginTransCompleteEvent += value;
			}
			remove
			{
				_BeginTransCompleteEvent -= value;
			}
		}

		/// <summary>
		/// SupportByVersion ADODB, 2.1,2.5
		/// </summary>
		private event Connection_CommitTransCompleteEventHandler _CommitTransCompleteEvent;

		/// <summary>
		/// SupportByVersion ADODB 2.1 2.5
		/// </summary>
		[SupportByVersion("ADODB", 2.1,2.5)]
		public event Connection_CommitTransCompleteEventHandler CommitTransCompleteEvent
		{
			add
			{
				CreateEventBridge();
				_CommitTransCompleteEvent += value;
			}
			remove
			{
				_CommitTransCompleteEvent -= value;
			}
		}

		/// <summary>
		/// SupportByVersion ADODB, 2.1,2.5
		/// </summary>
		private event Connection_RollbackTransCompleteEventHandler _RollbackTransCompleteEvent;

		/// <summary>
		/// SupportByVersion ADODB 2.1 2.5
		/// </summary>
		[SupportByVersion("ADODB", 2.1,2.5)]
		public event Connection_RollbackTransCompleteEventHandler RollbackTransCompleteEvent
		{
			add
			{
				CreateEventBridge();
				_RollbackTransCompleteEvent += value;
			}
			remove
			{
				_RollbackTransCompleteEvent -= value;
			}
		}

		/// <summary>
		/// SupportByVersion ADODB, 2.1,2.5
		/// </summary>
		private event Connection_WillExecuteEventHandler _WillExecuteEvent;

		/// <summary>
		/// SupportByVersion ADODB 2.1 2.5
		/// </summary>
		[SupportByVersion("ADODB", 2.1,2.5)]
		public event Connection_WillExecuteEventHandler WillExecuteEvent
		{
			add
			{
				CreateEventBridge();
				_WillExecuteEvent += value;
			}
			remove
			{
				_WillExecuteEvent -= value;
			}
		}

		/// <summary>
		/// SupportByVersion ADODB, 2.1,2.5
		/// </summary>
		private event Connection_ExecuteCompleteEventHandler _ExecuteCompleteEvent;

		/// <summary>
		/// SupportByVersion ADODB 2.1 2.5
		/// </summary>
		[SupportByVersion("ADODB", 2.1,2.5)]
		public event Connection_ExecuteCompleteEventHandler ExecuteCompleteEvent
		{
			add
			{
				CreateEventBridge();
				_ExecuteCompleteEvent += value;
			}
			remove
			{
				_ExecuteCompleteEvent -= value;
			}
		}

		/// <summary>
		/// SupportByVersion ADODB, 2.1,2.5
		/// </summary>
		private event Connection_WillConnectEventHandler _WillConnectEvent;

		/// <summary>
		/// SupportByVersion ADODB 2.1 2.5
		/// </summary>
		[SupportByVersion("ADODB", 2.1,2.5)]
		public event Connection_WillConnectEventHandler WillConnectEvent
		{
			add
			{
				CreateEventBridge();
				_WillConnectEvent += value;
			}
			remove
			{
				_WillConnectEvent -= value;
			}
		}

		/// <summary>
		/// SupportByVersion ADODB, 2.1,2.5
		/// </summary>
		private event Connection_ConnectCompleteEventHandler _ConnectCompleteEvent;

		/// <summary>
		/// SupportByVersion ADODB 2.1 2.5
		/// </summary>
		[SupportByVersion("ADODB", 2.1,2.5)]
		public event Connection_ConnectCompleteEventHandler ConnectCompleteEvent
		{
			add
			{
				CreateEventBridge();
				_ConnectCompleteEvent += value;
			}
			remove
			{
				_ConnectCompleteEvent -= value;
			}
		}

		/// <summary>
		/// SupportByVersion ADODB, 2.1,2.5
		/// </summary>
		private event Connection_DisconnectEventHandler _DisconnectEvent;

		/// <summary>
		/// SupportByVersion ADODB 2.1 2.5
		/// </summary>
		[SupportByVersion("ADODB", 2.1,2.5)]
		public event Connection_DisconnectEventHandler DisconnectEvent
		{
			add
			{
				CreateEventBridge();
				_DisconnectEvent += value;
			}
			remove
			{
				_DisconnectEvent -= value;
			}
		}

		#endregion
       
	    #region IEventBinding Member
        
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
				_activeSinkId = SinkHelper.GetConnectionPoint(this, ref _connectPoint, ConnectionEvents_SinkHelper.Id);


			if(ConnectionEvents_SinkHelper.Id.Equals(_activeSinkId, StringComparison.InvariantCultureIgnoreCase))
			{
				_connectionEvents_SinkHelper = new ConnectionEvents_SinkHelper(this, _connectPoint);
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
        ///  The instance has currently one or more event recipients 
        /// </summary>
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
        
        /// <summary>
        /// Target methods from its actual event recipients
        /// </summary>
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
       
        /// <summary>
        /// Returns the current count of event recipients
        /// </summary>
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
        
        /// <summary>
        /// Raise an instance event
        /// </summary>
        /// <param name="eventName">name of the event without 'Event' at the end</param>
        /// <param name="paramsArray">custom arguments for the event</param>
        /// <returns>count of called event recipients</returns>
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

        /// <summary>
        /// Stop listening events for the instance
        /// </summary>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        public void DisposeEventBridge()
        {
			if( null != _connectionEvents_SinkHelper)
			{
				_connectionEvents_SinkHelper.Dispose();
				_connectionEvents_SinkHelper = null;
			}

			_connectPoint = null;
		}
        
        #endregion

		#pragma warning restore
	}
}