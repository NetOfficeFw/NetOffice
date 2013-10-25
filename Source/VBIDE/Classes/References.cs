using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice;
namespace NetOffice.VBIDEApi
{

	#region Delegates

	#pragma warning disable
	public delegate void References_ItemAddedEventHandler(NetOffice.VBIDEApi.Reference Reference);
	public delegate void References_ItemRemovedEventHandler(NetOffice.VBIDEApi.Reference Reference);
	#pragma warning restore

	#endregion

	///<summary>
	/// CoClass References 
	/// SupportByVersion VBIDE, 12,14,5.3
	///</summary>
	[SupportByVersionAttribute("VBIDE", 12,14,5.3)]
	[EntityTypeAttribute(EntityType.IsCoClass)]
	public class References : _References,IEventBinding
	{
		#pragma warning disable
		#region Fields
		
		private NetRuntimeSystem.Runtime.InteropServices.ComTypes.IConnectionPoint _connectPoint;
		private string _activeSinkId;
		private NetRuntimeSystem.Type _thisType;
		_dispReferences_Events_SinkHelper __dispReferences_Events_SinkHelper;
	
		#endregion

		#region Type Information

        private static Type _type;
		
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        public static Type LateBindingApiWrapperType
        {
            get
            {
                if (null == _type)
                    _type = typeof(References);
                return _type;
            }
        }
        
        #endregion
        		
		#region Construction

		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
		public References(Core factory, COMObject parentObject, object comProxy) : base(factory, parentObject, comProxy)
		{
			
		}

        ///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
		public References(COMObject parentObject, object comProxy) : base(parentObject, comProxy)
		{
			
		}

		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public References(Core factory, COMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(factory, parentObject, comProxy, comProxyType)
		{
			
		}

		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public References(COMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(parentObject, comProxy, comProxyType)
		{
			
		}
		
		///<param name="replacedObject">object to replaced. replacedObject are not usable after this action</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public References(COMObject replacedObject) : base(replacedObject)
		{
			
		}
		
		///<summary>
        ///creates a new instance of References 
        ///</summary>		
		public References():base("VBIDE.References")
		{
			
		}
		
		///<summary>
        ///creates a new instance of References
        ///</summary>
        ///<param name="progId">registered ProgID</param>
		public References(string progId):base(progId)
		{
			
		}

		#endregion

		#region Static CoClass Methods

		/// <summary>
        /// returns all running VBIDE.References objects from the running object table(ROT)
        /// </summary>
        /// <returns>an VBIDE.References array</returns>
		public static NetOffice.VBIDEApi.References[] GetActiveInstances()
		{		
			NetRuntimeSystem.Collections.Generic.List<object> proxyList = NetOffice.RunningObjectTable.GetActiveProxiesFromROT("VBIDE","References");
			NetRuntimeSystem.Collections.Generic.List<NetOffice.VBIDEApi.References> resultList = new NetRuntimeSystem.Collections.Generic.List<NetOffice.VBIDEApi.References>();
			foreach(object proxy in proxyList)
				resultList.Add( new NetOffice.VBIDEApi.References(null, proxy) );
			return resultList.ToArray();
		}

		/// <summary>
        /// returns a running VBIDE.References object from the running object table(ROT). the method takes the first element from the table
        /// </summary>
        /// <returns>an VBIDE.References object or null</returns>
		public static NetOffice.VBIDEApi.References GetActiveInstance()
		{
			object proxy = NetOffice.RunningObjectTable.GetActiveProxyFromROT("VBIDE","References", false);
			if(null != proxy)
				return new NetOffice.VBIDEApi.References(null, proxy);
			else
				return null;
		}

		/// <summary>
        /// returns a running VBIDE.References object from the running object table(ROT).  the method takes the first element from the table
        /// </summary>
	    /// <param name="throwOnError">throw an exception if no object was found</param>
        /// <returns>an VBIDE.References object or null</returns>
		public static NetOffice.VBIDEApi.References GetActiveInstance(bool throwOnError)
		{
			object proxy = NetOffice.RunningObjectTable.GetActiveProxyFromROT("VBIDE","References", throwOnError);
			if(null != proxy)
				return new NetOffice.VBIDEApi.References(null, proxy);
			else
				return null;
		}
		#endregion

		#region Events

		/// <summary>
		/// SupportByVersion VBIDE, 12,14,5.3
		/// </summary>
		private event References_ItemAddedEventHandler _ItemAddedEvent;

		/// <summary>
		/// SupportByVersion VBIDE 12 14 5.3
		/// </summary>
		[SupportByVersion("VBIDE", 12,14,5.3)]
		public event References_ItemAddedEventHandler ItemAddedEvent
		{
			add
			{
				CreateEventBridge();
				_ItemAddedEvent += value;
			}
			remove
			{
				_ItemAddedEvent -= value;
			}
		}

		/// <summary>
		/// SupportByVersion VBIDE, 12,14,5.3
		/// </summary>
		private event References_ItemRemovedEventHandler _ItemRemovedEvent;

		/// <summary>
		/// SupportByVersion VBIDE 12 14 5.3
		/// </summary>
		[SupportByVersion("VBIDE", 12,14,5.3)]
		public event References_ItemRemovedEventHandler ItemRemovedEvent
		{
			add
			{
				CreateEventBridge();
				_ItemRemovedEvent += value;
			}
			remove
			{
				_ItemRemovedEvent -= value;
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
				_activeSinkId = SinkHelper.GetConnectionPoint(this, ref _connectPoint, _dispReferences_Events_SinkHelper.Id);


			if(_dispReferences_Events_SinkHelper.Id.Equals(_activeSinkId, StringComparison.InvariantCultureIgnoreCase))
			{
				__dispReferences_Events_SinkHelper = new _dispReferences_Events_SinkHelper(this, _connectPoint);
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
			if( null != __dispReferences_Events_SinkHelper)
			{
				__dispReferences_Events_SinkHelper.Dispose();
				__dispReferences_Events_SinkHelper = null;
			}

			_connectPoint = null;
		}
        
        #endregion

		#pragma warning restore
	}
}