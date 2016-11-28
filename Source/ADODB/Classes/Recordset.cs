using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice;
using NetOffice.Misc;

namespace NetOffice.ADODBApi
{

	#region Delegates

	#pragma warning disable
	public delegate void Recordset_WillChangeFieldEventHandler(Int32 cFields, object Fields, NetOffice.ADODBApi.Enums.EventStatusEnum adStatus, NetOffice.ADODBApi._Recordset pRecordset);
	public delegate void Recordset_FieldChangeCompleteEventHandler(Int32 cFields, object Fields, NetOffice.ADODBApi.Error pError, NetOffice.ADODBApi.Enums.EventStatusEnum adStatus, NetOffice.ADODBApi._Recordset pRecordset);
	public delegate void Recordset_WillChangeRecordEventHandler(NetOffice.ADODBApi.Enums.EventReasonEnum adReason, Int32 cRecords, NetOffice.ADODBApi.Enums.EventStatusEnum adStatus, NetOffice.ADODBApi._Recordset pRecordset);
	public delegate void Recordset_RecordChangeCompleteEventHandler(NetOffice.ADODBApi.Enums.EventReasonEnum adReason, Int32 cRecords, NetOffice.ADODBApi.Error pError, NetOffice.ADODBApi.Enums.EventStatusEnum adStatus, NetOffice.ADODBApi._Recordset pRecordset);
	public delegate void Recordset_WillChangeRecordsetEventHandler(NetOffice.ADODBApi.Enums.EventReasonEnum adReason, NetOffice.ADODBApi.Enums.EventStatusEnum adStatus, NetOffice.ADODBApi._Recordset pRecordset);
	public delegate void Recordset_RecordsetChangeCompleteEventHandler(NetOffice.ADODBApi.Enums.EventReasonEnum adReason, NetOffice.ADODBApi.Error pError, NetOffice.ADODBApi.Enums.EventStatusEnum adStatus, NetOffice.ADODBApi._Recordset pRecordset);
	public delegate void Recordset_WillMoveEventHandler(NetOffice.ADODBApi.Enums.EventReasonEnum adReason, NetOffice.ADODBApi.Enums.EventStatusEnum adStatus, NetOffice.ADODBApi._Recordset pRecordset);
	public delegate void Recordset_MoveCompleteEventHandler(NetOffice.ADODBApi.Enums.EventReasonEnum adReason, NetOffice.ADODBApi.Error pError, NetOffice.ADODBApi.Enums.EventStatusEnum adStatus, NetOffice.ADODBApi._Recordset pRecordset);
	public delegate void Recordset_EndOfRecordsetEventHandler(ref bool fMoreData, NetOffice.ADODBApi.Enums.EventStatusEnum adStatus, NetOffice.ADODBApi._Recordset pRecordset);
	public delegate void Recordset_FetchProgressEventHandler(Int32 Progress, Int32 MaxProgress, NetOffice.ADODBApi.Enums.EventStatusEnum adStatus, NetOffice.ADODBApi._Recordset pRecordset);
	public delegate void Recordset_FetchCompleteEventHandler(NetOffice.ADODBApi.Error pError, NetOffice.ADODBApi.Enums.EventStatusEnum adStatus, NetOffice.ADODBApi._Recordset pRecordset);
	#pragma warning restore

	#endregion

	///<summary>
	/// CoClass Recordset 
	/// SupportByVersion ADODB, 2.1,2.5
	///</summary>
	[SupportByVersionAttribute("ADODB", 2.1,2.5)]
	[EntityTypeAttribute(EntityType.IsCoClass)]
	public class Recordset : _Recordset,IEventBinding
	{
		#pragma warning disable
		#region Fields
		
		private NetRuntimeSystem.Runtime.InteropServices.ComTypes.IConnectionPoint _connectPoint;
		private string _activeSinkId;
		private NetRuntimeSystem.Type _thisType;
		RecordsetEvents_SinkHelper _recordsetEvents_SinkHelper;
	
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
                    _type = typeof(Recordset);
                return _type;
            }
        }
        
        #endregion
        		
		#region Construction

		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
		public Recordset(Core factory, ICOMObject parentObject, object comProxy) : base(factory, parentObject, comProxy)
		{
			
		}

        ///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
		public Recordset(ICOMObject parentObject, object comProxy) : base(parentObject, comProxy)
		{
			
		}

		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Recordset(Core factory, ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(factory, parentObject, comProxy, comProxyType)
		{
			
		}

		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Recordset(ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(parentObject, comProxy, comProxyType)
		{
			
		}
		
		///<param name="replacedObject">object to replaced. replacedObject are not usable after this action</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Recordset(ICOMObject replacedObject) : base(replacedObject)
		{
			
		}
		
		///<summary>
        /// Creates a new instance of Recordset 
        ///</summary>		
		public Recordset():base("ADODB.Recordset")
		{
			
		}
		
		///<summary>
        /// Creates a new instance of Recordset
        ///</summary>
        ///<param name="progId">registered ProgID</param>
		public Recordset(string progId):base(progId)
		{
			
		}

		#endregion

		#region Static CoClass Methods

		/// <summary>
        /// Returns all running ADODB.Recordset objects from the environment/system
        /// </summary>
        /// <returns>an ADODB.Recordset array</returns>
		public static NetOffice.ADODBApi.Recordset[] GetActiveInstances()
		{		
			IDisposableEnumeration proxyList = NetOffice.ProxyService.GetActiveInstances("ADODB","Recordset");
			NetRuntimeSystem.Collections.Generic.List<NetOffice.ADODBApi.Recordset> resultList = new NetRuntimeSystem.Collections.Generic.List<NetOffice.ADODBApi.Recordset>();
			foreach(object proxy in proxyList)
				resultList.Add( new NetOffice.ADODBApi.Recordset(null, proxy) );
			return resultList.ToArray();
		}

		/// <summary>
        /// Returns a running ADODB.Recordset object from the environment/system.
        /// </summary>
        /// <returns>an ADODB.Recordset object or null</returns>
		public static NetOffice.ADODBApi.Recordset GetActiveInstance()
		{
			object proxy  = NetOffice.ProxyService.GetActiveInstance("ADODB","Recordset", false);
			if(null != proxy)
				return new NetOffice.ADODBApi.Recordset(null, proxy);
			else
				return null;
		}

		/// <summary>
        /// Returns a running ADODB.Recordset object from the environment/system. 
        /// </summary>
	    /// <param name="throwOnError">throw an exception if no object was found</param>
        /// <returns>an ADODB.Recordset object or null</returns>
		public static NetOffice.ADODBApi.Recordset GetActiveInstance(bool throwOnError)
		{
			object proxy  = NetOffice.ProxyService.GetActiveInstance("ADODB","Recordset", throwOnError);
			if(null != proxy)
				return new NetOffice.ADODBApi.Recordset(null, proxy);
			else
				return null;
		}
		#endregion

		#region Events

		/// <summary>
		/// SupportByVersion ADODB, 2.1,2.5
		/// </summary>
		private event Recordset_WillChangeFieldEventHandler _WillChangeFieldEvent;

		/// <summary>
		/// SupportByVersion ADODB 2.1 2.5
		/// </summary>
		[SupportByVersion("ADODB", 2.1,2.5)]
		public event Recordset_WillChangeFieldEventHandler WillChangeFieldEvent
		{
			add
			{
				CreateEventBridge();
				_WillChangeFieldEvent += value;
			}
			remove
			{
				_WillChangeFieldEvent -= value;
			}
		}

		/// <summary>
		/// SupportByVersion ADODB, 2.1,2.5
		/// </summary>
		private event Recordset_FieldChangeCompleteEventHandler _FieldChangeCompleteEvent;

		/// <summary>
		/// SupportByVersion ADODB 2.1 2.5
		/// </summary>
		[SupportByVersion("ADODB", 2.1,2.5)]
		public event Recordset_FieldChangeCompleteEventHandler FieldChangeCompleteEvent
		{
			add
			{
				CreateEventBridge();
				_FieldChangeCompleteEvent += value;
			}
			remove
			{
				_FieldChangeCompleteEvent -= value;
			}
		}

		/// <summary>
		/// SupportByVersion ADODB, 2.1,2.5
		/// </summary>
		private event Recordset_WillChangeRecordEventHandler _WillChangeRecordEvent;

		/// <summary>
		/// SupportByVersion ADODB 2.1 2.5
		/// </summary>
		[SupportByVersion("ADODB", 2.1,2.5)]
		public event Recordset_WillChangeRecordEventHandler WillChangeRecordEvent
		{
			add
			{
				CreateEventBridge();
				_WillChangeRecordEvent += value;
			}
			remove
			{
				_WillChangeRecordEvent -= value;
			}
		}

		/// <summary>
		/// SupportByVersion ADODB, 2.1,2.5
		/// </summary>
		private event Recordset_RecordChangeCompleteEventHandler _RecordChangeCompleteEvent;

		/// <summary>
		/// SupportByVersion ADODB 2.1 2.5
		/// </summary>
		[SupportByVersion("ADODB", 2.1,2.5)]
		public event Recordset_RecordChangeCompleteEventHandler RecordChangeCompleteEvent
		{
			add
			{
				CreateEventBridge();
				_RecordChangeCompleteEvent += value;
			}
			remove
			{
				_RecordChangeCompleteEvent -= value;
			}
		}

		/// <summary>
		/// SupportByVersion ADODB, 2.1,2.5
		/// </summary>
		private event Recordset_WillChangeRecordsetEventHandler _WillChangeRecordsetEvent;

		/// <summary>
		/// SupportByVersion ADODB 2.1 2.5
		/// </summary>
		[SupportByVersion("ADODB", 2.1,2.5)]
		public event Recordset_WillChangeRecordsetEventHandler WillChangeRecordsetEvent
		{
			add
			{
				CreateEventBridge();
				_WillChangeRecordsetEvent += value;
			}
			remove
			{
				_WillChangeRecordsetEvent -= value;
			}
		}

		/// <summary>
		/// SupportByVersion ADODB, 2.1,2.5
		/// </summary>
		private event Recordset_RecordsetChangeCompleteEventHandler _RecordsetChangeCompleteEvent;

		/// <summary>
		/// SupportByVersion ADODB 2.1 2.5
		/// </summary>
		[SupportByVersion("ADODB", 2.1,2.5)]
		public event Recordset_RecordsetChangeCompleteEventHandler RecordsetChangeCompleteEvent
		{
			add
			{
				CreateEventBridge();
				_RecordsetChangeCompleteEvent += value;
			}
			remove
			{
				_RecordsetChangeCompleteEvent -= value;
			}
		}

		/// <summary>
		/// SupportByVersion ADODB, 2.1,2.5
		/// </summary>
		private event Recordset_WillMoveEventHandler _WillMoveEvent;

		/// <summary>
		/// SupportByVersion ADODB 2.1 2.5
		/// </summary>
		[SupportByVersion("ADODB", 2.1,2.5)]
		public event Recordset_WillMoveEventHandler WillMoveEvent
		{
			add
			{
				CreateEventBridge();
				_WillMoveEvent += value;
			}
			remove
			{
				_WillMoveEvent -= value;
			}
		}

		/// <summary>
		/// SupportByVersion ADODB, 2.1,2.5
		/// </summary>
		private event Recordset_MoveCompleteEventHandler _MoveCompleteEvent;

		/// <summary>
		/// SupportByVersion ADODB 2.1 2.5
		/// </summary>
		[SupportByVersion("ADODB", 2.1,2.5)]
		public event Recordset_MoveCompleteEventHandler MoveCompleteEvent
		{
			add
			{
				CreateEventBridge();
				_MoveCompleteEvent += value;
			}
			remove
			{
				_MoveCompleteEvent -= value;
			}
		}

		/// <summary>
		/// SupportByVersion ADODB, 2.1,2.5
		/// </summary>
		private event Recordset_EndOfRecordsetEventHandler _EndOfRecordsetEvent;

		/// <summary>
		/// SupportByVersion ADODB 2.1 2.5
		/// </summary>
		[SupportByVersion("ADODB", 2.1,2.5)]
		public event Recordset_EndOfRecordsetEventHandler EndOfRecordsetEvent
		{
			add
			{
				CreateEventBridge();
				_EndOfRecordsetEvent += value;
			}
			remove
			{
				_EndOfRecordsetEvent -= value;
			}
		}

		/// <summary>
		/// SupportByVersion ADODB, 2.1,2.5
		/// </summary>
		private event Recordset_FetchProgressEventHandler _FetchProgressEvent;

		/// <summary>
		/// SupportByVersion ADODB 2.1 2.5
		/// </summary>
		[SupportByVersion("ADODB", 2.1,2.5)]
		public event Recordset_FetchProgressEventHandler FetchProgressEvent
		{
			add
			{
				CreateEventBridge();
				_FetchProgressEvent += value;
			}
			remove
			{
				_FetchProgressEvent -= value;
			}
		}

		/// <summary>
		/// SupportByVersion ADODB, 2.1,2.5
		/// </summary>
		private event Recordset_FetchCompleteEventHandler _FetchCompleteEvent;

		/// <summary>
		/// SupportByVersion ADODB 2.1 2.5
		/// </summary>
		[SupportByVersion("ADODB", 2.1,2.5)]
		public event Recordset_FetchCompleteEventHandler FetchCompleteEvent
		{
			add
			{
				CreateEventBridge();
				_FetchCompleteEvent += value;
			}
			remove
			{
				_FetchCompleteEvent -= value;
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
				_activeSinkId = SinkHelper.GetConnectionPoint(this, ref _connectPoint, RecordsetEvents_SinkHelper.Id);


			if(RecordsetEvents_SinkHelper.Id.Equals(_activeSinkId, StringComparison.InvariantCultureIgnoreCase))
			{
				_recordsetEvents_SinkHelper = new RecordsetEvents_SinkHelper(this, _connectPoint);
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
			if( null != _recordsetEvents_SinkHelper)
			{
				_recordsetEvents_SinkHelper.Dispose();
				_recordsetEvents_SinkHelper = null;
			}

			_connectPoint = null;
		}
        
        #endregion

		#pragma warning restore
	}
}