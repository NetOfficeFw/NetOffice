using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice;
using NetOffice.Misc;

namespace NetOffice.OfficeApi
{

	#region Delegates

	#pragma warning disable
	public delegate void CustomXMLPart_NodeAfterInsertEventHandler(NetOffice.OfficeApi.CustomXMLNode NewNode, bool InUndoRedo);
	public delegate void CustomXMLPart_NodeAfterDeleteEventHandler(NetOffice.OfficeApi.CustomXMLNode OldNode, NetOffice.OfficeApi.CustomXMLNode OldParentNode, NetOffice.OfficeApi.CustomXMLNode OldNextSibling, bool InUndoRedo);
	public delegate void CustomXMLPart_NodeAfterReplaceEventHandler(NetOffice.OfficeApi.CustomXMLNode OldNode, NetOffice.OfficeApi.CustomXMLNode NewNode, bool InUndoRedo);
	#pragma warning restore

	#endregion

	///<summary>
	/// CoClass CustomXMLPart 
	/// SupportByVersion Office, 12,14,15,16
	/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff863497.aspx
	///</summary>
	[SupportByVersionAttribute("Office", 12,14,15,16)]
	[EntityTypeAttribute(EntityType.IsCoClass)]
	public class CustomXMLPart : _CustomXMLPart,IEventBinding
	{
		#pragma warning disable
		#region Fields
		
		private NetRuntimeSystem.Runtime.InteropServices.ComTypes.IConnectionPoint _connectPoint;
		private string _activeSinkId;
		private NetRuntimeSystem.Type _thisType;
		_CustomXMLPartEvents_SinkHelper __CustomXMLPartEvents_SinkHelper;
	
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
                    _type = typeof(CustomXMLPart);
                return _type;
            }
        }
        
        #endregion
        		
		#region Construction

		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
		public CustomXMLPart(Core factory, ICOMObject parentObject, object comProxy) : base(factory, parentObject, comProxy)
		{
			
		}

        ///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
		public CustomXMLPart(ICOMObject parentObject, object comProxy) : base(parentObject, comProxy)
		{
			
		}

		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public CustomXMLPart(Core factory, ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(factory, parentObject, comProxy, comProxyType)
		{
			
		}

		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public CustomXMLPart(ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(parentObject, comProxy, comProxyType)
		{
			
		}
		
		///<param name="replacedObject">object to replaced. replacedObject are not usable after this action</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public CustomXMLPart(ICOMObject replacedObject) : base(replacedObject)
		{
			
		}
		
		///<summary>
        /// Creates a new instance of CustomXMLPart 
        ///</summary>		
		public CustomXMLPart():base("Office.CustomXMLPart")
		{
			
		}
		
		///<summary>
        /// Creates a new instance of CustomXMLPart
        ///</summary>
        ///<param name="progId">registered ProgID</param>
		public CustomXMLPart(string progId):base(progId)
		{
			
		}

		#endregion

		#region Static CoClass Methods

		/// <summary>
        /// Returns all running Office.CustomXMLPart objects from the environment/system
        /// </summary>
        /// <returns>an Office.CustomXMLPart array</returns>
		public static NetOffice.OfficeApi.CustomXMLPart[] GetActiveInstances()
		{		
			IDisposableEnumeration proxyList = NetOffice.ProxyService.GetActiveInstances("Office","CustomXMLPart");
			NetRuntimeSystem.Collections.Generic.List<NetOffice.OfficeApi.CustomXMLPart> resultList = new NetRuntimeSystem.Collections.Generic.List<NetOffice.OfficeApi.CustomXMLPart>();
			foreach(object proxy in proxyList)
				resultList.Add( new NetOffice.OfficeApi.CustomXMLPart(null, proxy) );
			return resultList.ToArray();
		}

		/// <summary>
        /// Returns a running Office.CustomXMLPart object from the environment/system.
        /// </summary>
        /// <returns>an Office.CustomXMLPart object or null</returns>
		public static NetOffice.OfficeApi.CustomXMLPart GetActiveInstance()
		{
			object proxy  = NetOffice.ProxyService.GetActiveInstance("Office","CustomXMLPart", false);
			if(null != proxy)
				return new NetOffice.OfficeApi.CustomXMLPart(null, proxy);
			else
				return null;
		}

		/// <summary>
        /// Returns a running Office.CustomXMLPart object from the environment/system. 
        /// </summary>
	    /// <param name="throwOnError">throw an exception if no object was found</param>
        /// <returns>an Office.CustomXMLPart object or null</returns>
		public static NetOffice.OfficeApi.CustomXMLPart GetActiveInstance(bool throwOnError)
		{
			object proxy  = NetOffice.ProxyService.GetActiveInstance("Office","CustomXMLPart", throwOnError);
			if(null != proxy)
				return new NetOffice.OfficeApi.CustomXMLPart(null, proxy);
			else
				return null;
		}
		#endregion

		#region Events

		/// <summary>
		/// SupportByVersion Office, 12,14,15,16
		/// </summary>
		private event CustomXMLPart_NodeAfterInsertEventHandler _NodeAfterInsertEvent;

		/// <summary>
		/// SupportByVersion Office 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff862780.aspx </remarks>
		[SupportByVersion("Office", 12,14,15,16)]
		public event CustomXMLPart_NodeAfterInsertEventHandler NodeAfterInsertEvent
		{
			add
			{
				CreateEventBridge();
				_NodeAfterInsertEvent += value;
			}
			remove
			{
				_NodeAfterInsertEvent -= value;
			}
		}

		/// <summary>
		/// SupportByVersion Office, 12,14,15,16
		/// </summary>
		private event CustomXMLPart_NodeAfterDeleteEventHandler _NodeAfterDeleteEvent;

		/// <summary>
		/// SupportByVersion Office 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff861395.aspx </remarks>
		[SupportByVersion("Office", 12,14,15,16)]
		public event CustomXMLPart_NodeAfterDeleteEventHandler NodeAfterDeleteEvent
		{
			add
			{
				CreateEventBridge();
				_NodeAfterDeleteEvent += value;
			}
			remove
			{
				_NodeAfterDeleteEvent -= value;
			}
		}

		/// <summary>
		/// SupportByVersion Office, 12,14,15,16
		/// </summary>
		private event CustomXMLPart_NodeAfterReplaceEventHandler _NodeAfterReplaceEvent;

		/// <summary>
		/// SupportByVersion Office 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff863732.aspx </remarks>
		[SupportByVersion("Office", 12,14,15,16)]
		public event CustomXMLPart_NodeAfterReplaceEventHandler NodeAfterReplaceEvent
		{
			add
			{
				CreateEventBridge();
				_NodeAfterReplaceEvent += value;
			}
			remove
			{
				_NodeAfterReplaceEvent -= value;
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
				_activeSinkId = SinkHelper.GetConnectionPoint(this, ref _connectPoint, _CustomXMLPartEvents_SinkHelper.Id);


			if(_CustomXMLPartEvents_SinkHelper.Id.Equals(_activeSinkId, StringComparison.InvariantCultureIgnoreCase))
			{
				__CustomXMLPartEvents_SinkHelper = new _CustomXMLPartEvents_SinkHelper(this, _connectPoint);
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
			if( null != __CustomXMLPartEvents_SinkHelper)
			{
				__CustomXMLPartEvents_SinkHelper.Dispose();
				__CustomXMLPartEvents_SinkHelper = null;
			}

			_connectPoint = null;
		}
        
        #endregion

		#pragma warning restore
	}
}