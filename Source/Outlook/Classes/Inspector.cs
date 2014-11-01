using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice;
namespace NetOffice.OutlookApi
{

	#region Delegates

	#pragma warning disable
	public delegate void Inspector_ActivateEventHandler();
	public delegate void Inspector_DeactivateEventHandler();
	public delegate void Inspector_CloseEventHandler();
	public delegate void Inspector_BeforeMaximizeEventHandler(ref bool Cancel);
	public delegate void Inspector_BeforeMinimizeEventHandler(ref bool Cancel);
	public delegate void Inspector_BeforeMoveEventHandler(ref bool Cancel);
	public delegate void Inspector_BeforeSizeEventHandler(ref bool Cancel);
	public delegate void Inspector_PageChangeEventHandler(ref string ActivePageName);
	public delegate void Inspector_AttachmentSelectionChangeEventHandler();
	#pragma warning restore

	#endregion

	///<summary>
	/// CoClass Inspector 
	/// SupportByVersion Outlook, 9,10,11,12,14,15
	/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff869356.aspx
	///</summary>
	[SupportByVersionAttribute("Outlook", 9,10,11,12,14,15)]
	[EntityTypeAttribute(EntityType.IsCoClass)]
	public class Inspector : _Inspector,IEventBinding
	{
		#pragma warning disable
		#region Fields
		
		private NetRuntimeSystem.Runtime.InteropServices.ComTypes.IConnectionPoint _connectPoint;
		private string _activeSinkId;
		private NetRuntimeSystem.Type _thisType;
		InspectorEvents_SinkHelper _inspectorEvents_SinkHelper;
		InspectorEvents_10_SinkHelper _inspectorEvents_10_SinkHelper;
	
		#endregion

		#region Type Information

        private static Type _type;
		
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        public static Type LateBindingApiWrapperType
        {
            get
            {
                if (null == _type)
                    _type = typeof(Inspector);
                return _type;
            }
        }
        
        #endregion
        		
		#region Construction

		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
		public Inspector(Core factory, COMObject parentObject, object comProxy) : base(factory, parentObject, comProxy)
		{
			
		}

        ///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
		public Inspector(COMObject parentObject, object comProxy) : base(parentObject, comProxy)
		{
			
		}

		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Inspector(Core factory, COMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(factory, parentObject, comProxy, comProxyType)
		{
			
		}

		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Inspector(COMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(parentObject, comProxy, comProxyType)
		{
			
		}
		
		///<param name="replacedObject">object to replaced. replacedObject are not usable after this action</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Inspector(COMObject replacedObject) : base(replacedObject)
		{
			
		}
		
		///<summary>
        ///creates a new instance of Inspector 
        ///</summary>		
		public Inspector():base("Outlook.Inspector")
		{
			
		}
		
		///<summary>
        ///creates a new instance of Inspector
        ///</summary>
        ///<param name="progId">registered ProgID</param>
		public Inspector(string progId):base(progId)
		{
			
		}

		#endregion

		#region Static CoClass Methods

		/// <summary>
        /// returns all running Outlook.Inspector objects from the running object table(ROT)
        /// </summary>
        /// <returns>an Outlook.Inspector array</returns>
		public static NetOffice.OutlookApi.Inspector[] GetActiveInstances()
		{		
			NetRuntimeSystem.Collections.Generic.List<object> proxyList = NetOffice.RunningObjectTable.GetActiveProxiesFromROT("Outlook","Inspector");
			NetRuntimeSystem.Collections.Generic.List<NetOffice.OutlookApi.Inspector> resultList = new NetRuntimeSystem.Collections.Generic.List<NetOffice.OutlookApi.Inspector>();
			foreach(object proxy in proxyList)
				resultList.Add( new NetOffice.OutlookApi.Inspector(null, proxy) );
			return resultList.ToArray();
		}

		/// <summary>
        /// returns a running Outlook.Inspector object from the running object table(ROT). the method takes the first element from the table
        /// </summary>
        /// <returns>an Outlook.Inspector object or null</returns>
		public static NetOffice.OutlookApi.Inspector GetActiveInstance()
		{
			object proxy = NetOffice.RunningObjectTable.GetActiveProxyFromROT("Outlook","Inspector", false);
			if(null != proxy)
				return new NetOffice.OutlookApi.Inspector(null, proxy);
			else
				return null;
		}

		/// <summary>
        /// returns a running Outlook.Inspector object from the running object table(ROT).  the method takes the first element from the table
        /// </summary>
	    /// <param name="throwOnError">throw an exception if no object was found</param>
        /// <returns>an Outlook.Inspector object or null</returns>
		public static NetOffice.OutlookApi.Inspector GetActiveInstance(bool throwOnError)
		{
			object proxy = NetOffice.RunningObjectTable.GetActiveProxyFromROT("Outlook","Inspector", throwOnError);
			if(null != proxy)
				return new NetOffice.OutlookApi.Inspector(null, proxy);
			else
				return null;
		}
		#endregion

		#region Events

		/// <summary>
		/// SupportByVersion Outlook, 9,10,11,12,14,15
		/// </summary>
		private event Inspector_ActivateEventHandler _ActivateEvent;

		/// <summary>
		/// SupportByVersion Outlook 9 10 11 12 14 15
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff865363.aspx </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15)]
		public event Inspector_ActivateEventHandler ActivateEvent
		{
			add
			{
				CreateEventBridge();
				_ActivateEvent += value;
			}
			remove
			{
				_ActivateEvent -= value;
			}
		}

		/// <summary>
		/// SupportByVersion Outlook, 9,10,11,12,14,15
		/// </summary>
		private event Inspector_DeactivateEventHandler _DeactivateEvent;

		/// <summary>
		/// SupportByVersion Outlook 9 10 11 12 14 15
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff862214.aspx </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15)]
		public event Inspector_DeactivateEventHandler DeactivateEvent
		{
			add
			{
				CreateEventBridge();
				_DeactivateEvent += value;
			}
			remove
			{
				_DeactivateEvent -= value;
			}
		}

		/// <summary>
		/// SupportByVersion Outlook, 9,10,11,12,14,15
		/// </summary>
		private event Inspector_CloseEventHandler _CloseEvent;

		/// <summary>
		/// SupportByVersion Outlook 9 10 11 12 14 15
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff865374.aspx </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15)]
		public event Inspector_CloseEventHandler CloseEvent
		{
			add
			{
				CreateEventBridge();
				_CloseEvent += value;
			}
			remove
			{
				_CloseEvent -= value;
			}
		}

		/// <summary>
		/// SupportByVersion Outlook, 10,11,12,14,15
		/// </summary>
		private event Inspector_BeforeMaximizeEventHandler _BeforeMaximizeEvent;

		/// <summary>
		/// SupportByVersion Outlook 10 11 12 14 15
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff867903.aspx </remarks>
		[SupportByVersion("Outlook", 10,11,12,14,15)]
		public event Inspector_BeforeMaximizeEventHandler BeforeMaximizeEvent
		{
			add
			{
				CreateEventBridge();
				_BeforeMaximizeEvent += value;
			}
			remove
			{
				_BeforeMaximizeEvent -= value;
			}
		}

		/// <summary>
		/// SupportByVersion Outlook, 10,11,12,14,15
		/// </summary>
		private event Inspector_BeforeMinimizeEventHandler _BeforeMinimizeEvent;

		/// <summary>
		/// SupportByVersion Outlook 10 11 12 14 15
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff868289.aspx </remarks>
		[SupportByVersion("Outlook", 10,11,12,14,15)]
		public event Inspector_BeforeMinimizeEventHandler BeforeMinimizeEvent
		{
			add
			{
				CreateEventBridge();
				_BeforeMinimizeEvent += value;
			}
			remove
			{
				_BeforeMinimizeEvent -= value;
			}
		}

		/// <summary>
		/// SupportByVersion Outlook, 10,11,12,14,15
		/// </summary>
		private event Inspector_BeforeMoveEventHandler _BeforeMoveEvent;

		/// <summary>
		/// SupportByVersion Outlook 10 11 12 14 15
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff865042.aspx </remarks>
		[SupportByVersion("Outlook", 10,11,12,14,15)]
		public event Inspector_BeforeMoveEventHandler BeforeMoveEvent
		{
			add
			{
				CreateEventBridge();
				_BeforeMoveEvent += value;
			}
			remove
			{
				_BeforeMoveEvent -= value;
			}
		}

		/// <summary>
		/// SupportByVersion Outlook, 10,11,12,14,15
		/// </summary>
		private event Inspector_BeforeSizeEventHandler _BeforeSizeEvent;

		/// <summary>
		/// SupportByVersion Outlook 10 11 12 14 15
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff869786.aspx </remarks>
		[SupportByVersion("Outlook", 10,11,12,14,15)]
		public event Inspector_BeforeSizeEventHandler BeforeSizeEvent
		{
			add
			{
				CreateEventBridge();
				_BeforeSizeEvent += value;
			}
			remove
			{
				_BeforeSizeEvent -= value;
			}
		}

		/// <summary>
		/// SupportByVersion Outlook, 12,14,15
		/// </summary>
		private event Inspector_PageChangeEventHandler _PageChangeEvent;

		/// <summary>
		/// SupportByVersion Outlook 12 14 15
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff869845.aspx </remarks>
		[SupportByVersion("Outlook", 12,14,15)]
		public event Inspector_PageChangeEventHandler PageChangeEvent
		{
			add
			{
				CreateEventBridge();
				_PageChangeEvent += value;
			}
			remove
			{
				_PageChangeEvent -= value;
			}
		}

		/// <summary>
		/// SupportByVersion Outlook, 14,15
		/// </summary>
		private event Inspector_AttachmentSelectionChangeEventHandler _AttachmentSelectionChangeEvent;

		/// <summary>
		/// SupportByVersion Outlook 14 15
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff861296.aspx </remarks>
		[SupportByVersion("Outlook", 14,15)]
		public event Inspector_AttachmentSelectionChangeEventHandler AttachmentSelectionChangeEvent
		{
			add
			{
				CreateEventBridge();
				_AttachmentSelectionChangeEvent += value;
			}
			remove
			{
				_AttachmentSelectionChangeEvent -= value;
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
				_activeSinkId = SinkHelper.GetConnectionPoint(this, ref _connectPoint, InspectorEvents_SinkHelper.Id,InspectorEvents_10_SinkHelper.Id);


			if(InspectorEvents_SinkHelper.Id.Equals(_activeSinkId, StringComparison.InvariantCultureIgnoreCase))
			{
				_inspectorEvents_SinkHelper = new InspectorEvents_SinkHelper(this, _connectPoint);
				return;
			}

			if(InspectorEvents_10_SinkHelper.Id.Equals(_activeSinkId, StringComparison.InvariantCultureIgnoreCase))
			{
				_inspectorEvents_10_SinkHelper = new InspectorEvents_10_SinkHelper(this, _connectPoint);
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
			if( null != _inspectorEvents_SinkHelper)
			{
				_inspectorEvents_SinkHelper.Dispose();
				_inspectorEvents_SinkHelper = null;
			}
			if( null != _inspectorEvents_10_SinkHelper)
			{
				_inspectorEvents_10_SinkHelper.Dispose();
				_inspectorEvents_10_SinkHelper = null;
			}

			_connectPoint = null;
		}
        
        #endregion

		#pragma warning restore
	}
}