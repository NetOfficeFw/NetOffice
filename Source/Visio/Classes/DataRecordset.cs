using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice;
namespace NetOffice.VisioApi
{

	#region Delegates

	#pragma warning disable
	public delegate void DataRecordset_DataRecordsetChangedEventHandler(NetOffice.VisioApi.IVDataRecordsetChangedEvent DataRecordsetChanged);
	public delegate void DataRecordset_BeforeDataRecordsetDeleteEventHandler(NetOffice.VisioApi.IVDataRecordset DataRecordset);
	#pragma warning restore

	#endregion

	///<summary>
	/// CoClass DataRecordset 
	/// SupportByVersion Visio, 12,14,15
	/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/ff769258(v=office.14).aspx
	///</summary>
	[SupportByVersionAttribute("Visio", 12,14,15)]
	[EntityTypeAttribute(EntityType.IsCoClass)]
	public class DataRecordset : IVDataRecordset,IEventBinding
	{
		#pragma warning disable
		#region Fields
		
		private NetRuntimeSystem.Runtime.InteropServices.ComTypes.IConnectionPoint _connectPoint;
		private string _activeSinkId;
		private NetRuntimeSystem.Type _thisType;
		EDataRecordset_SinkHelper _eDataRecordset_SinkHelper;
	
		#endregion

		#region Type Information

        private static Type _type;
		
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        public static Type LateBindingApiWrapperType
        {
            get
            {
                if (null == _type)
                    _type = typeof(DataRecordset);
                return _type;
            }
        }
        
        #endregion
        		
		#region Construction

		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
		public DataRecordset(Core factory, COMObject parentObject, object comProxy) : base(factory, parentObject, comProxy)
		{
			
		}

        ///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
		public DataRecordset(COMObject parentObject, object comProxy) : base(parentObject, comProxy)
		{
			
		}

		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public DataRecordset(Core factory, COMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(factory, parentObject, comProxy, comProxyType)
		{
			
		}

		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public DataRecordset(COMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(parentObject, comProxy, comProxyType)
		{
			
		}
		
		///<param name="replacedObject">object to replaced. replacedObject are not usable after this action</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public DataRecordset(COMObject replacedObject) : base(replacedObject)
		{
			
		}
		
		///<summary>
        ///creates a new instance of DataRecordset 
        ///</summary>		
		public DataRecordset():base("Visio.DataRecordset")
		{
			
		}
		
		///<summary>
        ///creates a new instance of DataRecordset
        ///</summary>
        ///<param name="progId">registered ProgID</param>
		public DataRecordset(string progId):base(progId)
		{
			
		}

		#endregion

		#region Static CoClass Methods

		/// <summary>
        /// returns all running Visio.DataRecordset objects from the running object table(ROT)
        /// </summary>
        /// <returns>an Visio.DataRecordset array</returns>
		public static NetOffice.VisioApi.DataRecordset[] GetActiveInstances()
		{		
			NetRuntimeSystem.Collections.Generic.List<object> proxyList = NetOffice.RunningObjectTable.GetActiveProxiesFromROT("Visio","DataRecordset");
			NetRuntimeSystem.Collections.Generic.List<NetOffice.VisioApi.DataRecordset> resultList = new NetRuntimeSystem.Collections.Generic.List<NetOffice.VisioApi.DataRecordset>();
			foreach(object proxy in proxyList)
				resultList.Add( new NetOffice.VisioApi.DataRecordset(null, proxy) );
			return resultList.ToArray();
		}

		/// <summary>
        /// returns a running Visio.DataRecordset object from the running object table(ROT). the method takes the first element from the table
        /// </summary>
        /// <returns>an Visio.DataRecordset object or null</returns>
		public static NetOffice.VisioApi.DataRecordset GetActiveInstance()
		{
			object proxy = NetOffice.RunningObjectTable.GetActiveProxyFromROT("Visio","DataRecordset", false);
			if(null != proxy)
				return new NetOffice.VisioApi.DataRecordset(null, proxy);
			else
				return null;
		}

		/// <summary>
        /// returns a running Visio.DataRecordset object from the running object table(ROT).  the method takes the first element from the table
        /// </summary>
	    /// <param name="throwOnError">throw an exception if no object was found</param>
        /// <returns>an Visio.DataRecordset object or null</returns>
		public static NetOffice.VisioApi.DataRecordset GetActiveInstance(bool throwOnError)
		{
			object proxy = NetOffice.RunningObjectTable.GetActiveProxyFromROT("Visio","DataRecordset", throwOnError);
			if(null != proxy)
				return new NetOffice.VisioApi.DataRecordset(null, proxy);
			else
				return null;
		}
		#endregion

		#region Events

		/// <summary>
		/// SupportByVersion Visio, 12,14,15
		/// </summary>
		private event DataRecordset_DataRecordsetChangedEventHandler _DataRecordsetChangedEvent;

		/// <summary>
		/// SupportByVersion Visio 12 14 15
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/ff766153(v=office.14).aspx </remarks>
		[SupportByVersion("Visio", 12,14,15)]
		public event DataRecordset_DataRecordsetChangedEventHandler DataRecordsetChangedEvent
		{
			add
			{
				CreateEventBridge();
				_DataRecordsetChangedEvent += value;
			}
			remove
			{
				_DataRecordsetChangedEvent -= value;
			}
		}

		/// <summary>
		/// SupportByVersion Visio, 12,14,15
		/// </summary>
		private event DataRecordset_BeforeDataRecordsetDeleteEventHandler _BeforeDataRecordsetDeleteEvent;

		/// <summary>
		/// SupportByVersion Visio 12 14 15
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/ff766834(v=office.14).aspx </remarks>
		[SupportByVersion("Visio", 12,14,15)]
		public event DataRecordset_BeforeDataRecordsetDeleteEventHandler BeforeDataRecordsetDeleteEvent
		{
			add
			{
				CreateEventBridge();
				_BeforeDataRecordsetDeleteEvent += value;
			}
			remove
			{
				_BeforeDataRecordsetDeleteEvent -= value;
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
				_activeSinkId = SinkHelper.GetConnectionPoint(this, ref _connectPoint, EDataRecordset_SinkHelper.Id);


			if(EDataRecordset_SinkHelper.Id.Equals(_activeSinkId, StringComparison.InvariantCultureIgnoreCase))
			{
				_eDataRecordset_SinkHelper = new EDataRecordset_SinkHelper(this, _connectPoint);
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
			if( null != _eDataRecordset_SinkHelper)
			{
				_eDataRecordset_SinkHelper.Dispose();
				_eDataRecordset_SinkHelper = null;
			}

			_connectPoint = null;
		}
        
        #endregion

		#pragma warning restore
	}
}