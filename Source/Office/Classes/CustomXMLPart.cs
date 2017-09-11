using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.OfficeApi
{
	#region Delegates

	#pragma warning disable
	public delegate void CustomXMLPart_NodeAfterInsertEventHandler(NetOffice.OfficeApi.CustomXMLNode newNode, bool InUndoRedo);
	public delegate void CustomXMLPart_NodeAfterDeleteEventHandler(NetOffice.OfficeApi.CustomXMLNode oldNode, NetOffice.OfficeApi.CustomXMLNode oldParentNode, NetOffice.OfficeApi.CustomXMLNode oldNextSibling, bool inUndoRedo);
	public delegate void CustomXMLPart_NodeAfterReplaceEventHandler(NetOffice.OfficeApi.CustomXMLNode oldNode, NetOffice.OfficeApi.CustomXMLNode newNode, bool inUndoRedo);
	#pragma warning restore

	#endregion

	/// <summary>
	/// CoClass CustomXMLPart 
	/// SupportByVersion Office, 12,14,15,16
	/// </summary>
	/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff863497.aspx </remarks>
	[SupportByVersion("Office", 12,14,15,16)]
	[EntityType(EntityType.IsCoClass)]
	[EventSink(typeof(Events._CustomXMLPartEvents_SinkHelper))]
    [ComEventInterface(typeof(Events._CustomXMLPartEvents))]
    public class CustomXMLPart : _CustomXMLPart, IEventBinding
	{
		#pragma warning disable

		#region Fields
		
		private NetRuntimeSystem.Runtime.InteropServices.ComTypes.IConnectionPoint _connectPoint;
		private string _activeSinkId;
        private static Type _type;
        private Events._CustomXMLPartEvents_SinkHelper __CustomXMLPartEvents_SinkHelper;
	
		#endregion

		#region Type Information

        /// <summary>
        /// Instance Type
        /// </summary>
		[EditorBrowsable(EditorBrowsableState.Advanced), Browsable(false), Category("NetOffice"), CoreOverridden]
        public override Type InstanceType
        {
            get
            {
                return LateBindingApiWrapperType;
            }
        }

        /// <summary>
        /// Type Cache
        /// </summary>
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
		
		/// <summary>
        /// Creates a new instance of CustomXMLPart 
        /// </summary>		
		public CustomXMLPart():base("Office.CustomXMLPart")
		{
			
		}
		
		/// <summary>
        /// Creates a new instance of CustomXMLPart
        /// </summary>
        ///<param name="progId">registered ProgID</param>
		public CustomXMLPart(string progId):base(progId)
		{
			
		}

		#endregion

		#region Static CoClass Methods
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
       
	    #region IEventBinding
        
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
				_activeSinkId = SinkHelper.GetConnectionPoint(this, ref _connectPoint, Events._CustomXMLPartEvents_SinkHelper.Id);


			if(Events._CustomXMLPartEvents_SinkHelper.Id.Equals(_activeSinkId, StringComparison.InvariantCultureIgnoreCase))
			{
				__CustomXMLPartEvents_SinkHelper = new Events._CustomXMLPartEvents_SinkHelper(this, _connectPoint);
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
        /// Instance has one or more event recipients
        /// </summary>
        /// <returns>true if one or more event is active, otherwise false</returns>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        public bool HasEventRecipients()       
        {
            return NetOffice.Events.CoClassEventReflector.HasEventRecipients(this, LateBindingApiWrapperType);            
        }

        /// <summary>
        /// Instance has one or more event recipients
        /// </summary>
        /// <param name="eventName">name of the event</param>
        /// <returns></returns>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        public bool HasEventRecipients(string eventName)
        {
            return NetOffice.Events.CoClassEventReflector.HasEventRecipients(this, LateBindingApiWrapperType, eventName);
        }

        /// <summary>
        /// Target methods from its actual event recipients
        /// </summary>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        public Delegate[] GetEventRecipients(string eventName)
        {
            return NetOffice.Events.CoClassEventReflector.GetEventRecipients(this, LateBindingApiWrapperType, eventName);
        }
       
        /// <summary>
        /// Returns the current count of event recipients
        /// </summary>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        public int GetCountOfEventRecipients(string eventName)
        {
            return NetOffice.Events.CoClassEventReflector.GetCountOfEventRecipients(this, LateBindingApiWrapperType, eventName);       
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
            return NetOffice.Events.CoClassEventReflector.RaiseCustomEvent(this, LateBindingApiWrapperType, eventName, ref paramsArray);
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

