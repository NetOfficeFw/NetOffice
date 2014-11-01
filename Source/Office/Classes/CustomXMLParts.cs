using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice;
namespace NetOffice.OfficeApi
{

	#region Delegates

	#pragma warning disable
	public delegate void CustomXMLParts_PartAfterAddEventHandler(NetOffice.OfficeApi.CustomXMLPart NewPart);
	public delegate void CustomXMLParts_PartBeforeDeleteEventHandler(NetOffice.OfficeApi.CustomXMLPart OldPart);
	public delegate void CustomXMLParts_PartAfterLoadEventHandler(NetOffice.OfficeApi.CustomXMLPart Part);
	#pragma warning restore

	#endregion

	///<summary>
	/// CoClass CustomXMLParts 
	/// SupportByVersion Office, 12,14,15
	/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff863162.aspx
	///</summary>
	[SupportByVersionAttribute("Office", 12,14,15)]
	[EntityTypeAttribute(EntityType.IsCoClass)]
	public class CustomXMLParts : _CustomXMLParts,IEventBinding
	{
		#pragma warning disable
		#region Fields
		
		private NetRuntimeSystem.Runtime.InteropServices.ComTypes.IConnectionPoint _connectPoint;
		private string _activeSinkId;
		private NetRuntimeSystem.Type _thisType;
		_CustomXMLPartsEvents_SinkHelper __CustomXMLPartsEvents_SinkHelper;
	
		#endregion

		#region Type Information

        private static Type _type;
		
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        public static Type LateBindingApiWrapperType
        {
            get
            {
                if (null == _type)
                    _type = typeof(CustomXMLParts);
                return _type;
            }
        }
        
        #endregion
        		
		#region Construction

		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
		public CustomXMLParts(Core factory, COMObject parentObject, object comProxy) : base(factory, parentObject, comProxy)
		{
			
		}

        ///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
		public CustomXMLParts(COMObject parentObject, object comProxy) : base(parentObject, comProxy)
		{
			
		}

		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public CustomXMLParts(Core factory, COMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(factory, parentObject, comProxy, comProxyType)
		{
			
		}

		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public CustomXMLParts(COMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(parentObject, comProxy, comProxyType)
		{
			
		}
		
		///<param name="replacedObject">object to replaced. replacedObject are not usable after this action</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public CustomXMLParts(COMObject replacedObject) : base(replacedObject)
		{
			
		}
		
		///<summary>
        ///creates a new instance of CustomXMLParts 
        ///</summary>		
		public CustomXMLParts():base("Office.CustomXMLParts")
		{
			
		}
		
		///<summary>
        ///creates a new instance of CustomXMLParts
        ///</summary>
        ///<param name="progId">registered ProgID</param>
		public CustomXMLParts(string progId):base(progId)
		{
			
		}

		#endregion

		#region Static CoClass Methods

		/// <summary>
        /// returns all running Office.CustomXMLParts objects from the running object table(ROT)
        /// </summary>
        /// <returns>an Office.CustomXMLParts array</returns>
		public static NetOffice.OfficeApi.CustomXMLParts[] GetActiveInstances()
		{		
			NetRuntimeSystem.Collections.Generic.List<object> proxyList = NetOffice.RunningObjectTable.GetActiveProxiesFromROT("Office","CustomXMLParts");
			NetRuntimeSystem.Collections.Generic.List<NetOffice.OfficeApi.CustomXMLParts> resultList = new NetRuntimeSystem.Collections.Generic.List<NetOffice.OfficeApi.CustomXMLParts>();
			foreach(object proxy in proxyList)
				resultList.Add( new NetOffice.OfficeApi.CustomXMLParts(null, proxy) );
			return resultList.ToArray();
		}

		/// <summary>
        /// returns a running Office.CustomXMLParts object from the running object table(ROT). the method takes the first element from the table
        /// </summary>
        /// <returns>an Office.CustomXMLParts object or null</returns>
		public static NetOffice.OfficeApi.CustomXMLParts GetActiveInstance()
		{
			object proxy = NetOffice.RunningObjectTable.GetActiveProxyFromROT("Office","CustomXMLParts", false);
			if(null != proxy)
				return new NetOffice.OfficeApi.CustomXMLParts(null, proxy);
			else
				return null;
		}

		/// <summary>
        /// returns a running Office.CustomXMLParts object from the running object table(ROT).  the method takes the first element from the table
        /// </summary>
	    /// <param name="throwOnError">throw an exception if no object was found</param>
        /// <returns>an Office.CustomXMLParts object or null</returns>
		public static NetOffice.OfficeApi.CustomXMLParts GetActiveInstance(bool throwOnError)
		{
			object proxy = NetOffice.RunningObjectTable.GetActiveProxyFromROT("Office","CustomXMLParts", throwOnError);
			if(null != proxy)
				return new NetOffice.OfficeApi.CustomXMLParts(null, proxy);
			else
				return null;
		}
		#endregion

		#region Events

		/// <summary>
		/// SupportByVersion Office, 12,14,15
		/// </summary>
		private event CustomXMLParts_PartAfterAddEventHandler _PartAfterAddEvent;

		/// <summary>
		/// SupportByVersion Office 12 14 15
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff864147.aspx </remarks>
		[SupportByVersion("Office", 12,14,15)]
		public event CustomXMLParts_PartAfterAddEventHandler PartAfterAddEvent
		{
			add
			{
				CreateEventBridge();
				_PartAfterAddEvent += value;
			}
			remove
			{
				_PartAfterAddEvent -= value;
			}
		}

		/// <summary>
		/// SupportByVersion Office, 12,14,15
		/// </summary>
		private event CustomXMLParts_PartBeforeDeleteEventHandler _PartBeforeDeleteEvent;

		/// <summary>
		/// SupportByVersion Office 12 14 15
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff861735.aspx </remarks>
		[SupportByVersion("Office", 12,14,15)]
		public event CustomXMLParts_PartBeforeDeleteEventHandler PartBeforeDeleteEvent
		{
			add
			{
				CreateEventBridge();
				_PartBeforeDeleteEvent += value;
			}
			remove
			{
				_PartBeforeDeleteEvent -= value;
			}
		}

		/// <summary>
		/// SupportByVersion Office, 12,14,15
		/// </summary>
		private event CustomXMLParts_PartAfterLoadEventHandler _PartAfterLoadEvent;

		/// <summary>
		/// SupportByVersion Office 12 14 15
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff864879.aspx </remarks>
		[SupportByVersion("Office", 12,14,15)]
		public event CustomXMLParts_PartAfterLoadEventHandler PartAfterLoadEvent
		{
			add
			{
				CreateEventBridge();
				_PartAfterLoadEvent += value;
			}
			remove
			{
				_PartAfterLoadEvent -= value;
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
				_activeSinkId = SinkHelper.GetConnectionPoint(this, ref _connectPoint, _CustomXMLPartsEvents_SinkHelper.Id);


			if(_CustomXMLPartsEvents_SinkHelper.Id.Equals(_activeSinkId, StringComparison.InvariantCultureIgnoreCase))
			{
				__CustomXMLPartsEvents_SinkHelper = new _CustomXMLPartsEvents_SinkHelper(this, _connectPoint);
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
			if( null != __CustomXMLPartsEvents_SinkHelper)
			{
				__CustomXMLPartsEvents_SinkHelper.Dispose();
				__CustomXMLPartsEvents_SinkHelper = null;
			}

			_connectPoint = null;
		}
        
        #endregion

		#pragma warning restore
	}
}