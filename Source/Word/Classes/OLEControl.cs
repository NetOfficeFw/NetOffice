using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice;
namespace NetOffice.WordApi
{

	#region Delegates

	#pragma warning disable
	public delegate void OLEControl_GotFocusEventHandler();
	public delegate void OLEControl_LostFocusEventHandler();
	#pragma warning restore

	#endregion

	///<summary>
	/// CoClass OLEControl 
	/// SupportByVersion Word, 9,10,11,12,14,15
	///</summary>
	[SupportByVersionAttribute("Word", 9,10,11,12,14,15)]
	[EntityTypeAttribute(EntityType.IsCoClass)]
	public class OLEControl : _OLEControl,IEventBinding
	{
		#pragma warning disable
		#region Fields
		
		private NetRuntimeSystem.Runtime.InteropServices.ComTypes.IConnectionPoint _connectPoint;
		private string _activeSinkId;
		private NetRuntimeSystem.Type _thisType;
		OCXEvents_SinkHelper _oCXEvents_SinkHelper;
	
		#endregion

		#region Type Information

        private static Type _type;
		
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        public static Type LateBindingApiWrapperType
        {
            get
            {
                if (null == _type)
                    _type = typeof(OLEControl);
                return _type;
            }
        }
        
        #endregion
        		
		#region Construction

		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
		public OLEControl(Core factory, COMObject parentObject, object comProxy) : base(factory, parentObject, comProxy)
		{
			
		}

        ///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
		public OLEControl(COMObject parentObject, object comProxy) : base(parentObject, comProxy)
		{
			
		}

		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public OLEControl(Core factory, COMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(factory, parentObject, comProxy, comProxyType)
		{
			
		}

		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public OLEControl(COMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(parentObject, comProxy, comProxyType)
		{
			
		}
		
		///<param name="replacedObject">object to replaced. replacedObject are not usable after this action</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public OLEControl(COMObject replacedObject) : base(replacedObject)
		{
			
		}
		
		///<summary>
        ///creates a new instance of OLEControl 
        ///</summary>		
		public OLEControl():base("Word.OLEControl")
		{
			
		}
		
		///<summary>
        ///creates a new instance of OLEControl
        ///</summary>
        ///<param name="progId">registered ProgID</param>
		public OLEControl(string progId):base(progId)
		{
			
		}

		#endregion

		#region Static CoClass Methods

		/// <summary>
        /// returns all running Word.OLEControl objects from the running object table(ROT)
        /// </summary>
        /// <returns>an Word.OLEControl array</returns>
		public static NetOffice.WordApi.OLEControl[] GetActiveInstances()
		{		
			NetRuntimeSystem.Collections.Generic.List<object> proxyList = NetOffice.RunningObjectTable.GetActiveProxiesFromROT("Word","OLEControl");
			NetRuntimeSystem.Collections.Generic.List<NetOffice.WordApi.OLEControl> resultList = new NetRuntimeSystem.Collections.Generic.List<NetOffice.WordApi.OLEControl>();
			foreach(object proxy in proxyList)
				resultList.Add( new NetOffice.WordApi.OLEControl(null, proxy) );
			return resultList.ToArray();
		}

		/// <summary>
        /// returns a running Word.OLEControl object from the running object table(ROT). the method takes the first element from the table
        /// </summary>
        /// <returns>an Word.OLEControl object or null</returns>
		public static NetOffice.WordApi.OLEControl GetActiveInstance()
		{
			object proxy = NetOffice.RunningObjectTable.GetActiveProxyFromROT("Word","OLEControl", false);
			if(null != proxy)
				return new NetOffice.WordApi.OLEControl(null, proxy);
			else
				return null;
		}

		/// <summary>
        /// returns a running Word.OLEControl object from the running object table(ROT).  the method takes the first element from the table
        /// </summary>
	    /// <param name="throwOnError">throw an exception if no object was found</param>
        /// <returns>an Word.OLEControl object or null</returns>
		public static NetOffice.WordApi.OLEControl GetActiveInstance(bool throwOnError)
		{
			object proxy = NetOffice.RunningObjectTable.GetActiveProxyFromROT("Word","OLEControl", throwOnError);
			if(null != proxy)
				return new NetOffice.WordApi.OLEControl(null, proxy);
			else
				return null;
		}
		#endregion

		#region Events

		/// <summary>
		/// SupportByVersion Word, 9,10,11,12,14,15
		/// </summary>
		private event OLEControl_GotFocusEventHandler _GotFocusEvent;

		/// <summary>
		/// SupportByVersion Word 9 10 11 12 14 15
		/// </summary>
		[SupportByVersion("Word", 9,10,11,12,14,15)]
		public event OLEControl_GotFocusEventHandler GotFocusEvent
		{
			add
			{
				CreateEventBridge();
				_GotFocusEvent += value;
			}
			remove
			{
				_GotFocusEvent -= value;
			}
		}

		/// <summary>
		/// SupportByVersion Word, 9,10,11,12,14,15
		/// </summary>
		private event OLEControl_LostFocusEventHandler _LostFocusEvent;

		/// <summary>
		/// SupportByVersion Word 9 10 11 12 14 15
		/// </summary>
		[SupportByVersion("Word", 9,10,11,12,14,15)]
		public event OLEControl_LostFocusEventHandler LostFocusEvent
		{
			add
			{
				CreateEventBridge();
				_LostFocusEvent += value;
			}
			remove
			{
				_LostFocusEvent -= value;
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
				_activeSinkId = SinkHelper.GetConnectionPoint(this, ref _connectPoint, OCXEvents_SinkHelper.Id);


			if(OCXEvents_SinkHelper.Id.Equals(_activeSinkId, StringComparison.InvariantCultureIgnoreCase))
			{
				_oCXEvents_SinkHelper = new OCXEvents_SinkHelper(this, _connectPoint);
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
			if( null != _oCXEvents_SinkHelper)
			{
				_oCXEvents_SinkHelper.Dispose();
				_oCXEvents_SinkHelper = null;
			}

			_connectPoint = null;
		}
        
        #endregion

		#pragma warning restore
	}
}