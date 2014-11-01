using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice;
namespace NetOffice.OutlookApi
{

	#region Delegates

	#pragma warning disable
	public delegate void OutlookBarGroups_GroupAddEventHandler(NetOffice.OutlookApi.OutlookBarGroup NewGroup);
	public delegate void OutlookBarGroups_BeforeGroupAddEventHandler(ref bool Cancel);
	public delegate void OutlookBarGroups_BeforeGroupRemoveEventHandler(NetOffice.OutlookApi.OutlookBarGroup Group, ref bool Cancel);
	#pragma warning restore

	#endregion

	///<summary>
	/// CoClass OutlookBarGroups 
	/// SupportByVersion Outlook, 9,10,11,12,14,15
	/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff868789.aspx
	///</summary>
	[SupportByVersionAttribute("Outlook", 9,10,11,12,14,15)]
	[EntityTypeAttribute(EntityType.IsCoClass)]
	public class OutlookBarGroups : _OutlookBarGroups,IEventBinding
	{
		#pragma warning disable
		#region Fields
		
		private NetRuntimeSystem.Runtime.InteropServices.ComTypes.IConnectionPoint _connectPoint;
		private string _activeSinkId;
		private NetRuntimeSystem.Type _thisType;
		OutlookBarGroupsEvents_SinkHelper _outlookBarGroupsEvents_SinkHelper;
	
		#endregion

		#region Type Information

        private static Type _type;
		
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        public static Type LateBindingApiWrapperType
        {
            get
            {
                if (null == _type)
                    _type = typeof(OutlookBarGroups);
                return _type;
            }
        }
        
        #endregion
        		
		#region Construction

		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
		public OutlookBarGroups(Core factory, COMObject parentObject, object comProxy) : base(factory, parentObject, comProxy)
		{
			
		}

        ///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
		public OutlookBarGroups(COMObject parentObject, object comProxy) : base(parentObject, comProxy)
		{
			
		}

		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public OutlookBarGroups(Core factory, COMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(factory, parentObject, comProxy, comProxyType)
		{
			
		}

		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public OutlookBarGroups(COMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(parentObject, comProxy, comProxyType)
		{
			
		}
		
		///<param name="replacedObject">object to replaced. replacedObject are not usable after this action</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public OutlookBarGroups(COMObject replacedObject) : base(replacedObject)
		{
			
		}
		
		///<summary>
        ///creates a new instance of OutlookBarGroups 
        ///</summary>		
		public OutlookBarGroups():base("Outlook.OutlookBarGroups")
		{
			
		}
		
		///<summary>
        ///creates a new instance of OutlookBarGroups
        ///</summary>
        ///<param name="progId">registered ProgID</param>
		public OutlookBarGroups(string progId):base(progId)
		{
			
		}

		#endregion

		#region Static CoClass Methods

		/// <summary>
        /// returns all running Outlook.OutlookBarGroups objects from the running object table(ROT)
        /// </summary>
        /// <returns>an Outlook.OutlookBarGroups array</returns>
		public static NetOffice.OutlookApi.OutlookBarGroups[] GetActiveInstances()
		{		
			NetRuntimeSystem.Collections.Generic.List<object> proxyList = NetOffice.RunningObjectTable.GetActiveProxiesFromROT("Outlook","OutlookBarGroups");
			NetRuntimeSystem.Collections.Generic.List<NetOffice.OutlookApi.OutlookBarGroups> resultList = new NetRuntimeSystem.Collections.Generic.List<NetOffice.OutlookApi.OutlookBarGroups>();
			foreach(object proxy in proxyList)
				resultList.Add( new NetOffice.OutlookApi.OutlookBarGroups(null, proxy) );
			return resultList.ToArray();
		}

		/// <summary>
        /// returns a running Outlook.OutlookBarGroups object from the running object table(ROT). the method takes the first element from the table
        /// </summary>
        /// <returns>an Outlook.OutlookBarGroups object or null</returns>
		public static NetOffice.OutlookApi.OutlookBarGroups GetActiveInstance()
		{
			object proxy = NetOffice.RunningObjectTable.GetActiveProxyFromROT("Outlook","OutlookBarGroups", false);
			if(null != proxy)
				return new NetOffice.OutlookApi.OutlookBarGroups(null, proxy);
			else
				return null;
		}

		/// <summary>
        /// returns a running Outlook.OutlookBarGroups object from the running object table(ROT).  the method takes the first element from the table
        /// </summary>
	    /// <param name="throwOnError">throw an exception if no object was found</param>
        /// <returns>an Outlook.OutlookBarGroups object or null</returns>
		public static NetOffice.OutlookApi.OutlookBarGroups GetActiveInstance(bool throwOnError)
		{
			object proxy = NetOffice.RunningObjectTable.GetActiveProxyFromROT("Outlook","OutlookBarGroups", throwOnError);
			if(null != proxy)
				return new NetOffice.OutlookApi.OutlookBarGroups(null, proxy);
			else
				return null;
		}
		#endregion

		#region Events

		/// <summary>
		/// SupportByVersion Outlook, 9,10,11,12,14,15
		/// </summary>
		private event OutlookBarGroups_GroupAddEventHandler _GroupAddEvent;

		/// <summary>
		/// SupportByVersion Outlook 9 10 11 12 14 15
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff865659.aspx </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15)]
		public event OutlookBarGroups_GroupAddEventHandler GroupAddEvent
		{
			add
			{
				CreateEventBridge();
				_GroupAddEvent += value;
			}
			remove
			{
				_GroupAddEvent -= value;
			}
		}

		/// <summary>
		/// SupportByVersion Outlook, 9,10,11,12,14,15
		/// </summary>
		private event OutlookBarGroups_BeforeGroupAddEventHandler _BeforeGroupAddEvent;

		/// <summary>
		/// SupportByVersion Outlook 9 10 11 12 14 15
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff866940.aspx </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15)]
		public event OutlookBarGroups_BeforeGroupAddEventHandler BeforeGroupAddEvent
		{
			add
			{
				CreateEventBridge();
				_BeforeGroupAddEvent += value;
			}
			remove
			{
				_BeforeGroupAddEvent -= value;
			}
		}

		/// <summary>
		/// SupportByVersion Outlook, 9,10,11,12,14,15
		/// </summary>
		private event OutlookBarGroups_BeforeGroupRemoveEventHandler _BeforeGroupRemoveEvent;

		/// <summary>
		/// SupportByVersion Outlook 9 10 11 12 14 15
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff868646.aspx </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15)]
		public event OutlookBarGroups_BeforeGroupRemoveEventHandler BeforeGroupRemoveEvent
		{
			add
			{
				CreateEventBridge();
				_BeforeGroupRemoveEvent += value;
			}
			remove
			{
				_BeforeGroupRemoveEvent -= value;
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
				_activeSinkId = SinkHelper.GetConnectionPoint(this, ref _connectPoint, OutlookBarGroupsEvents_SinkHelper.Id);


			if(OutlookBarGroupsEvents_SinkHelper.Id.Equals(_activeSinkId, StringComparison.InvariantCultureIgnoreCase))
			{
				_outlookBarGroupsEvents_SinkHelper = new OutlookBarGroupsEvents_SinkHelper(this, _connectPoint);
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
			if( null != _outlookBarGroupsEvents_SinkHelper)
			{
				_outlookBarGroupsEvents_SinkHelper.Dispose();
				_outlookBarGroupsEvents_SinkHelper = null;
			}

			_connectPoint = null;
		}
        
        #endregion

		#pragma warning restore
	}
}