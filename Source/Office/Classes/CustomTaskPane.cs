using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice;
namespace NetOffice.OfficeApi
{

	#region Delegates

	#pragma warning disable
	public delegate void CustomTaskPane_VisibleStateChangeEventHandler(NetOffice.OfficeApi._CustomTaskPane CustomTaskPaneInst);
	public delegate void CustomTaskPane_DockPositionStateChangeEventHandler(NetOffice.OfficeApi._CustomTaskPane CustomTaskPaneInst);
	#pragma warning restore

	#endregion

	///<summary>
	/// CoClass CustomTaskPane 
	/// SupportByVersion Office, 12,14,15
	/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff862782.aspx
	///</summary>
	[SupportByVersionAttribute("Office", 12,14,15)]
	[EntityTypeAttribute(EntityType.IsCoClass)]
	public class CustomTaskPane : _CustomTaskPane,IEventBinding
	{
		#pragma warning disable
		#region Fields
		
		private NetRuntimeSystem.Runtime.InteropServices.ComTypes.IConnectionPoint _connectPoint;
		private string _activeSinkId;
		private NetRuntimeSystem.Type _thisType;
		_CustomTaskPaneEvents_SinkHelper __CustomTaskPaneEvents_SinkHelper;
	
		#endregion

		#region Type Information

        private static Type _type;
		
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        public static Type LateBindingApiWrapperType
        {
            get
            {
                if (null == _type)
                    _type = typeof(CustomTaskPane);
                return _type;
            }
        }
        
        #endregion
        		
		#region Construction

		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
		public CustomTaskPane(Core factory, COMObject parentObject, object comProxy) : base(factory, parentObject, comProxy)
		{
			
		}

        ///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
		public CustomTaskPane(COMObject parentObject, object comProxy) : base(parentObject, comProxy)
		{
			
		}

		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public CustomTaskPane(Core factory, COMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(factory, parentObject, comProxy, comProxyType)
		{
			
		}

		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public CustomTaskPane(COMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(parentObject, comProxy, comProxyType)
		{
			
		}
		
		///<param name="replacedObject">object to replaced. replacedObject are not usable after this action</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public CustomTaskPane(COMObject replacedObject) : base(replacedObject)
		{
			
		}
		
		///<summary>
        ///creates a new instance of CustomTaskPane 
        ///</summary>		
		public CustomTaskPane():base("Office.CustomTaskPane")
		{
			
		}
		
		///<summary>
        ///creates a new instance of CustomTaskPane
        ///</summary>
        ///<param name="progId">registered ProgID</param>
		public CustomTaskPane(string progId):base(progId)
		{
			
		}

		#endregion

		#region Static CoClass Methods

		/// <summary>
        /// returns all running Office.CustomTaskPane objects from the running object table(ROT)
        /// </summary>
        /// <returns>an Office.CustomTaskPane array</returns>
		public static NetOffice.OfficeApi.CustomTaskPane[] GetActiveInstances()
		{		
			NetRuntimeSystem.Collections.Generic.List<object> proxyList = NetOffice.RunningObjectTable.GetActiveProxiesFromROT("Office","CustomTaskPane");
			NetRuntimeSystem.Collections.Generic.List<NetOffice.OfficeApi.CustomTaskPane> resultList = new NetRuntimeSystem.Collections.Generic.List<NetOffice.OfficeApi.CustomTaskPane>();
			foreach(object proxy in proxyList)
				resultList.Add( new NetOffice.OfficeApi.CustomTaskPane(null, proxy) );
			return resultList.ToArray();
		}

		/// <summary>
        /// returns a running Office.CustomTaskPane object from the running object table(ROT). the method takes the first element from the table
        /// </summary>
        /// <returns>an Office.CustomTaskPane object or null</returns>
		public static NetOffice.OfficeApi.CustomTaskPane GetActiveInstance()
		{
			object proxy = NetOffice.RunningObjectTable.GetActiveProxyFromROT("Office","CustomTaskPane", false);
			if(null != proxy)
				return new NetOffice.OfficeApi.CustomTaskPane(null, proxy);
			else
				return null;
		}

		/// <summary>
        /// returns a running Office.CustomTaskPane object from the running object table(ROT).  the method takes the first element from the table
        /// </summary>
	    /// <param name="throwOnError">throw an exception if no object was found</param>
        /// <returns>an Office.CustomTaskPane object or null</returns>
		public static NetOffice.OfficeApi.CustomTaskPane GetActiveInstance(bool throwOnError)
		{
			object proxy = NetOffice.RunningObjectTable.GetActiveProxyFromROT("Office","CustomTaskPane", throwOnError);
			if(null != proxy)
				return new NetOffice.OfficeApi.CustomTaskPane(null, proxy);
			else
				return null;
		}
		#endregion

		#region Events

		/// <summary>
		/// SupportByVersion Office, 12,14,15
		/// </summary>
		private event CustomTaskPane_VisibleStateChangeEventHandler _VisibleStateChangeEvent;

		/// <summary>
		/// SupportByVersion Office 12 14 15
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff862422.aspx </remarks>
		[SupportByVersion("Office", 12,14,15)]
		public event CustomTaskPane_VisibleStateChangeEventHandler VisibleStateChangeEvent
		{
			add
			{
				CreateEventBridge();
				_VisibleStateChangeEvent += value;
			}
			remove
			{
				_VisibleStateChangeEvent -= value;
			}
		}

		/// <summary>
		/// SupportByVersion Office, 12,14,15
		/// </summary>
		private event CustomTaskPane_DockPositionStateChangeEventHandler _DockPositionStateChangeEvent;

		/// <summary>
		/// SupportByVersion Office 12 14 15
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff865561.aspx </remarks>
		[SupportByVersion("Office", 12,14,15)]
		public event CustomTaskPane_DockPositionStateChangeEventHandler DockPositionStateChangeEvent
		{
			add
			{
				CreateEventBridge();
				_DockPositionStateChangeEvent += value;
			}
			remove
			{
				_DockPositionStateChangeEvent -= value;
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
				_activeSinkId = SinkHelper.GetConnectionPoint(this, ref _connectPoint, _CustomTaskPaneEvents_SinkHelper.Id);


			if(_CustomTaskPaneEvents_SinkHelper.Id.Equals(_activeSinkId, StringComparison.InvariantCultureIgnoreCase))
			{
				__CustomTaskPaneEvents_SinkHelper = new _CustomTaskPaneEvents_SinkHelper(this, _connectPoint);
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
			if( null != __CustomTaskPaneEvents_SinkHelper)
			{
				__CustomTaskPaneEvents_SinkHelper.Dispose();
				__CustomTaskPaneEvents_SinkHelper = null;
			}

			_connectPoint = null;
		}
        
        #endregion

		#pragma warning restore
	}
}