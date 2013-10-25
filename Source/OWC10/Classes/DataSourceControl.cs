using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice;
namespace NetOffice.OWC10Api
{

	#region Delegates

	#pragma warning disable
	public delegate void DataSourceControl_CurrentEventHandler(NetOffice.OWC10Api.DSCEventInfo DSCEventInfo);
	public delegate void DataSourceControl_BeforeExpandEventHandler(NetOffice.OWC10Api.DSCEventInfo DSCEventInfo);
	public delegate void DataSourceControl_BeforeCollapseEventHandler(NetOffice.OWC10Api.DSCEventInfo DSCEventInfo);
	public delegate void DataSourceControl_BeforeFirstPageEventHandler(NetOffice.OWC10Api.DSCEventInfo DSCEventInfo);
	public delegate void DataSourceControl_BeforePreviousPageEventHandler(NetOffice.OWC10Api.DSCEventInfo DSCEventInfo);
	public delegate void DataSourceControl_BeforeNextPageEventHandler(NetOffice.OWC10Api.DSCEventInfo DSCEventInfo);
	public delegate void DataSourceControl_BeforeLastPageEventHandler(NetOffice.OWC10Api.DSCEventInfo DSCEventInfo);
	public delegate void DataSourceControl_DataErrorEventHandler(NetOffice.OWC10Api.DSCEventInfo DSCEventInfo);
	public delegate void DataSourceControl_DataPageCompleteEventHandler(NetOffice.OWC10Api.DSCEventInfo DSCEventInfo);
	public delegate void DataSourceControl_BeforeInitialBindEventHandler(NetOffice.OWC10Api.DSCEventInfo DSCEventInfo);
	public delegate void DataSourceControl_RecordsetSaveProgressEventHandler(NetOffice.OWC10Api.DSCEventInfo DSCEventInfo);
	public delegate void DataSourceControl_AfterDeleteEventHandler(NetOffice.OWC10Api.DSCEventInfo DSCEventInfo);
	public delegate void DataSourceControl_AfterInsertEventHandler(NetOffice.OWC10Api.DSCEventInfo DSCEventInfo);
	public delegate void DataSourceControl_AfterUpdateEventHandler(NetOffice.OWC10Api.DSCEventInfo DSCEventInfo);
	public delegate void DataSourceControl_BeforeDeleteEventHandler(NetOffice.OWC10Api.DSCEventInfo DSCEventInfo);
	public delegate void DataSourceControl_BeforeInsertEventHandler(NetOffice.OWC10Api.DSCEventInfo DSCEventInfo);
	public delegate void DataSourceControl_BeforeOverwriteEventHandler(NetOffice.OWC10Api.DSCEventInfo DSCEventInfo);
	public delegate void DataSourceControl_BeforeUpdateEventHandler(NetOffice.OWC10Api.DSCEventInfo DSCEventInfo);
	public delegate void DataSourceControl_DirtyEventHandler(NetOffice.OWC10Api.DSCEventInfo DSCEventInfo);
	public delegate void DataSourceControl_RecordExitEventHandler(NetOffice.OWC10Api.DSCEventInfo DSCEventInfo);
	public delegate void DataSourceControl_UndoEventHandler(NetOffice.OWC10Api.DSCEventInfo DSCEventInfo);
	public delegate void DataSourceControl_FocusEventHandler(NetOffice.OWC10Api.DSCEventInfo DSCEventInfo);
	#pragma warning restore

	#endregion

	///<summary>
	/// CoClass DataSourceControl 
	/// SupportByVersion OWC10, 1
	///</summary>
	[SupportByVersionAttribute("OWC10", 1)]
	[EntityTypeAttribute(EntityType.IsCoClass)]
	public class DataSourceControl : IDataSourceControl,IEventBinding
	{
		#pragma warning disable
		#region Fields
		
		private NetRuntimeSystem.Runtime.InteropServices.ComTypes.IConnectionPoint _connectPoint;
		private string _activeSinkId;
		private NetRuntimeSystem.Type _thisType;
		_DataSourceControlEvent_SinkHelper __DataSourceControlEvent_SinkHelper;
	
		#endregion

		#region Type Information

        private static Type _type;
		
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        public static Type LateBindingApiWrapperType
        {
            get
            {
                if (null == _type)
                    _type = typeof(DataSourceControl);
                return _type;
            }
        }
        
        #endregion
        		
		#region Construction

		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
		public DataSourceControl(Core factory, COMObject parentObject, object comProxy) : base(factory, parentObject, comProxy)
		{
			
		}

        ///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
		public DataSourceControl(COMObject parentObject, object comProxy) : base(parentObject, comProxy)
		{
			
		}

		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public DataSourceControl(Core factory, COMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(factory, parentObject, comProxy, comProxyType)
		{
			
		}

		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public DataSourceControl(COMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(parentObject, comProxy, comProxyType)
		{
			
		}
		
		///<param name="replacedObject">object to replaced. replacedObject are not usable after this action</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public DataSourceControl(COMObject replacedObject) : base(replacedObject)
		{
			
		}
		
		///<summary>
        ///creates a new instance of DataSourceControl 
        ///</summary>		
		public DataSourceControl():base("OWC10.DataSourceControl")
		{
			
		}
		
		///<summary>
        ///creates a new instance of DataSourceControl
        ///</summary>
        ///<param name="progId">registered ProgID</param>
		public DataSourceControl(string progId):base(progId)
		{
			
		}

		#endregion

		#region Static CoClass Methods

		/// <summary>
        /// returns all running OWC10.DataSourceControl objects from the running object table(ROT)
        /// </summary>
        /// <returns>an OWC10.DataSourceControl array</returns>
		public static NetOffice.OWC10Api.DataSourceControl[] GetActiveInstances()
		{		
			NetRuntimeSystem.Collections.Generic.List<object> proxyList = NetOffice.RunningObjectTable.GetActiveProxiesFromROT("OWC10","DataSourceControl");
			NetRuntimeSystem.Collections.Generic.List<NetOffice.OWC10Api.DataSourceControl> resultList = new NetRuntimeSystem.Collections.Generic.List<NetOffice.OWC10Api.DataSourceControl>();
			foreach(object proxy in proxyList)
				resultList.Add( new NetOffice.OWC10Api.DataSourceControl(null, proxy) );
			return resultList.ToArray();
		}

		/// <summary>
        /// returns a running OWC10.DataSourceControl object from the running object table(ROT). the method takes the first element from the table
        /// </summary>
        /// <returns>an OWC10.DataSourceControl object or null</returns>
		public static NetOffice.OWC10Api.DataSourceControl GetActiveInstance()
		{
			object proxy = NetOffice.RunningObjectTable.GetActiveProxyFromROT("OWC10","DataSourceControl", false);
			if(null != proxy)
				return new NetOffice.OWC10Api.DataSourceControl(null, proxy);
			else
				return null;
		}

		/// <summary>
        /// returns a running OWC10.DataSourceControl object from the running object table(ROT).  the method takes the first element from the table
        /// </summary>
	    /// <param name="throwOnError">throw an exception if no object was found</param>
        /// <returns>an OWC10.DataSourceControl object or null</returns>
		public static NetOffice.OWC10Api.DataSourceControl GetActiveInstance(bool throwOnError)
		{
			object proxy = NetOffice.RunningObjectTable.GetActiveProxyFromROT("OWC10","DataSourceControl", throwOnError);
			if(null != proxy)
				return new NetOffice.OWC10Api.DataSourceControl(null, proxy);
			else
				return null;
		}
		#endregion

		#region Events

		/// <summary>
		/// SupportByVersion OWC10, 1
		/// </summary>
		private event DataSourceControl_CurrentEventHandler _CurrentEvent;

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public event DataSourceControl_CurrentEventHandler CurrentEvent
		{
			add
			{
				CreateEventBridge();
				_CurrentEvent += value;
			}
			remove
			{
				_CurrentEvent -= value;
			}
		}

		/// <summary>
		/// SupportByVersion OWC10, 1
		/// </summary>
		private event DataSourceControl_BeforeExpandEventHandler _BeforeExpandEvent;

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public event DataSourceControl_BeforeExpandEventHandler BeforeExpandEvent
		{
			add
			{
				CreateEventBridge();
				_BeforeExpandEvent += value;
			}
			remove
			{
				_BeforeExpandEvent -= value;
			}
		}

		/// <summary>
		/// SupportByVersion OWC10, 1
		/// </summary>
		private event DataSourceControl_BeforeCollapseEventHandler _BeforeCollapseEvent;

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public event DataSourceControl_BeforeCollapseEventHandler BeforeCollapseEvent
		{
			add
			{
				CreateEventBridge();
				_BeforeCollapseEvent += value;
			}
			remove
			{
				_BeforeCollapseEvent -= value;
			}
		}

		/// <summary>
		/// SupportByVersion OWC10, 1
		/// </summary>
		private event DataSourceControl_BeforeFirstPageEventHandler _BeforeFirstPageEvent;

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public event DataSourceControl_BeforeFirstPageEventHandler BeforeFirstPageEvent
		{
			add
			{
				CreateEventBridge();
				_BeforeFirstPageEvent += value;
			}
			remove
			{
				_BeforeFirstPageEvent -= value;
			}
		}

		/// <summary>
		/// SupportByVersion OWC10, 1
		/// </summary>
		private event DataSourceControl_BeforePreviousPageEventHandler _BeforePreviousPageEvent;

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public event DataSourceControl_BeforePreviousPageEventHandler BeforePreviousPageEvent
		{
			add
			{
				CreateEventBridge();
				_BeforePreviousPageEvent += value;
			}
			remove
			{
				_BeforePreviousPageEvent -= value;
			}
		}

		/// <summary>
		/// SupportByVersion OWC10, 1
		/// </summary>
		private event DataSourceControl_BeforeNextPageEventHandler _BeforeNextPageEvent;

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public event DataSourceControl_BeforeNextPageEventHandler BeforeNextPageEvent
		{
			add
			{
				CreateEventBridge();
				_BeforeNextPageEvent += value;
			}
			remove
			{
				_BeforeNextPageEvent -= value;
			}
		}

		/// <summary>
		/// SupportByVersion OWC10, 1
		/// </summary>
		private event DataSourceControl_BeforeLastPageEventHandler _BeforeLastPageEvent;

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public event DataSourceControl_BeforeLastPageEventHandler BeforeLastPageEvent
		{
			add
			{
				CreateEventBridge();
				_BeforeLastPageEvent += value;
			}
			remove
			{
				_BeforeLastPageEvent -= value;
			}
		}

		/// <summary>
		/// SupportByVersion OWC10, 1
		/// </summary>
		private event DataSourceControl_DataErrorEventHandler _DataErrorEvent;

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public event DataSourceControl_DataErrorEventHandler DataErrorEvent
		{
			add
			{
				CreateEventBridge();
				_DataErrorEvent += value;
			}
			remove
			{
				_DataErrorEvent -= value;
			}
		}

		/// <summary>
		/// SupportByVersion OWC10, 1
		/// </summary>
		private event DataSourceControl_DataPageCompleteEventHandler _DataPageCompleteEvent;

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public event DataSourceControl_DataPageCompleteEventHandler DataPageCompleteEvent
		{
			add
			{
				CreateEventBridge();
				_DataPageCompleteEvent += value;
			}
			remove
			{
				_DataPageCompleteEvent -= value;
			}
		}

		/// <summary>
		/// SupportByVersion OWC10, 1
		/// </summary>
		private event DataSourceControl_BeforeInitialBindEventHandler _BeforeInitialBindEvent;

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public event DataSourceControl_BeforeInitialBindEventHandler BeforeInitialBindEvent
		{
			add
			{
				CreateEventBridge();
				_BeforeInitialBindEvent += value;
			}
			remove
			{
				_BeforeInitialBindEvent -= value;
			}
		}

		/// <summary>
		/// SupportByVersion OWC10, 1
		/// </summary>
		private event DataSourceControl_RecordsetSaveProgressEventHandler _RecordsetSaveProgressEvent;

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public event DataSourceControl_RecordsetSaveProgressEventHandler RecordsetSaveProgressEvent
		{
			add
			{
				CreateEventBridge();
				_RecordsetSaveProgressEvent += value;
			}
			remove
			{
				_RecordsetSaveProgressEvent -= value;
			}
		}

		/// <summary>
		/// SupportByVersion OWC10, 1
		/// </summary>
		private event DataSourceControl_AfterDeleteEventHandler _AfterDeleteEvent;

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public event DataSourceControl_AfterDeleteEventHandler AfterDeleteEvent
		{
			add
			{
				CreateEventBridge();
				_AfterDeleteEvent += value;
			}
			remove
			{
				_AfterDeleteEvent -= value;
			}
		}

		/// <summary>
		/// SupportByVersion OWC10, 1
		/// </summary>
		private event DataSourceControl_AfterInsertEventHandler _AfterInsertEvent;

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public event DataSourceControl_AfterInsertEventHandler AfterInsertEvent
		{
			add
			{
				CreateEventBridge();
				_AfterInsertEvent += value;
			}
			remove
			{
				_AfterInsertEvent -= value;
			}
		}

		/// <summary>
		/// SupportByVersion OWC10, 1
		/// </summary>
		private event DataSourceControl_AfterUpdateEventHandler _AfterUpdateEvent;

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public event DataSourceControl_AfterUpdateEventHandler AfterUpdateEvent
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
		/// SupportByVersion OWC10, 1
		/// </summary>
		private event DataSourceControl_BeforeDeleteEventHandler _BeforeDeleteEvent;

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public event DataSourceControl_BeforeDeleteEventHandler BeforeDeleteEvent
		{
			add
			{
				CreateEventBridge();
				_BeforeDeleteEvent += value;
			}
			remove
			{
				_BeforeDeleteEvent -= value;
			}
		}

		/// <summary>
		/// SupportByVersion OWC10, 1
		/// </summary>
		private event DataSourceControl_BeforeInsertEventHandler _BeforeInsertEvent;

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public event DataSourceControl_BeforeInsertEventHandler BeforeInsertEvent
		{
			add
			{
				CreateEventBridge();
				_BeforeInsertEvent += value;
			}
			remove
			{
				_BeforeInsertEvent -= value;
			}
		}

		/// <summary>
		/// SupportByVersion OWC10, 1
		/// </summary>
		private event DataSourceControl_BeforeOverwriteEventHandler _BeforeOverwriteEvent;

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public event DataSourceControl_BeforeOverwriteEventHandler BeforeOverwriteEvent
		{
			add
			{
				CreateEventBridge();
				_BeforeOverwriteEvent += value;
			}
			remove
			{
				_BeforeOverwriteEvent -= value;
			}
		}

		/// <summary>
		/// SupportByVersion OWC10, 1
		/// </summary>
		private event DataSourceControl_BeforeUpdateEventHandler _BeforeUpdateEvent;

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public event DataSourceControl_BeforeUpdateEventHandler BeforeUpdateEvent
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
		/// SupportByVersion OWC10, 1
		/// </summary>
		private event DataSourceControl_DirtyEventHandler _DirtyEvent;

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public event DataSourceControl_DirtyEventHandler DirtyEvent
		{
			add
			{
				CreateEventBridge();
				_DirtyEvent += value;
			}
			remove
			{
				_DirtyEvent -= value;
			}
		}

		/// <summary>
		/// SupportByVersion OWC10, 1
		/// </summary>
		private event DataSourceControl_RecordExitEventHandler _RecordExitEvent;

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public event DataSourceControl_RecordExitEventHandler RecordExitEvent
		{
			add
			{
				CreateEventBridge();
				_RecordExitEvent += value;
			}
			remove
			{
				_RecordExitEvent -= value;
			}
		}

		/// <summary>
		/// SupportByVersion OWC10, 1
		/// </summary>
		private event DataSourceControl_UndoEventHandler _UndoEvent;

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public event DataSourceControl_UndoEventHandler UndoEvent
		{
			add
			{
				CreateEventBridge();
				_UndoEvent += value;
			}
			remove
			{
				_UndoEvent -= value;
			}
		}

		/// <summary>
		/// SupportByVersion OWC10, 1
		/// </summary>
		private event DataSourceControl_FocusEventHandler _FocusEvent;

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public event DataSourceControl_FocusEventHandler FocusEvent
		{
			add
			{
				CreateEventBridge();
				_FocusEvent += value;
			}
			remove
			{
				_FocusEvent -= value;
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
				_activeSinkId = SinkHelper.GetConnectionPoint(this, ref _connectPoint, _DataSourceControlEvent_SinkHelper.Id);


			if(_DataSourceControlEvent_SinkHelper.Id.Equals(_activeSinkId, StringComparison.InvariantCultureIgnoreCase))
			{
				__DataSourceControlEvent_SinkHelper = new _DataSourceControlEvent_SinkHelper(this, _connectPoint);
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
			if( null != __DataSourceControlEvent_SinkHelper)
			{
				__DataSourceControlEvent_SinkHelper.Dispose();
				__DataSourceControlEvent_SinkHelper = null;
			}

			_connectPoint = null;
		}
        
        #endregion

		#pragma warning restore
	}
}