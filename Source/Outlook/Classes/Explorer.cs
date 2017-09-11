using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.OutlookApi
{
	#region Delegates

	#pragma warning disable
	public delegate void Explorer_ActivateEventHandler();
	public delegate void Explorer_FolderSwitchEventHandler();
	public delegate void Explorer_BeforeFolderSwitchEventHandler(ICOMObject newFolder, ref bool cancel);
	public delegate void Explorer_ViewSwitchEventHandler();
	public delegate void Explorer_BeforeViewSwitchEventHandler(object newView, ref bool cancel);
	public delegate void Explorer_DeactivateEventHandler();
	public delegate void Explorer_SelectionChangeEventHandler();
	public delegate void Explorer_CloseEventHandler();
	public delegate void Explorer_BeforeMaximizeEventHandler(ref bool cancel);
	public delegate void Explorer_BeforeMinimizeEventHandler(ref bool cancel);
	public delegate void Explorer_BeforeMoveEventHandler(ref bool cancel);
	public delegate void Explorer_BeforeSizeEventHandler(ref bool cancel);
	public delegate void Explorer_BeforeItemCopyEventHandler(ref bool cancel);
	public delegate void Explorer_BeforeItemCutEventHandler(ref bool cancel);
	public delegate void Explorer_BeforeItemPasteEventHandler(ref object clipboardContent, NetOffice.OutlookApi.MAPIFolder target, ref bool cancel);
	public delegate void Explorer_AttachmentSelectionChangeEventHandler();
	public delegate void Explorer_InlineResponseEventHandler(ICOMObject item);
    public delegate void Explorer_InlineResponseCloseEventHandler();
	#pragma warning restore

	#endregion

	/// <summary>
	/// CoClass Explorer 
	/// SupportByVersion Outlook, 9,10,11,12,14,15,16
	/// </summary>
	/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff860356.aspx </remarks>
	[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
	[EntityType(EntityType.IsCoClass)]
	[EventSink(typeof(Events.ExplorerEvents_SinkHelper), typeof(Events.ExplorerEvents_10_SinkHelper))]
    [ComEventInterface(typeof(Events.ExplorerEvents), typeof(Events.ExplorerEvents_10))]
    public class Explorer : _Explorer, IEventBinding
	{
		#pragma warning disable

		#region Fields
		
		private NetRuntimeSystem.Runtime.InteropServices.ComTypes.IConnectionPoint _connectPoint;
		private string _activeSinkId;
        private static Type _type;
        private Events.ExplorerEvents_SinkHelper _explorerEvents_SinkHelper;
        private Events.ExplorerEvents_10_SinkHelper _explorerEvents_10_SinkHelper;
	
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
                    _type = typeof(Explorer);
                return _type;
            }
        }
        
        #endregion
        		
		#region Construction

		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
		public Explorer(Core factory, ICOMObject parentObject, object comProxy) : base(factory, parentObject, comProxy)
		{
			
		}

        ///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
		public Explorer(ICOMObject parentObject, object comProxy) : base(parentObject, comProxy)
		{
			
		}

		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Explorer(Core factory, ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(factory, parentObject, comProxy, comProxyType)
		{
			
		}

		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Explorer(ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(parentObject, comProxy, comProxyType)
		{
			
		}
		
		///<param name="replacedObject">object to replaced. replacedObject are not usable after this action</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Explorer(ICOMObject replacedObject) : base(replacedObject)
		{
			
		}
		
		/// <summary>
        /// Creates a new instance of Explorer 
        /// </summary>		
		public Explorer():base("Outlook.Explorer")
		{
			
		}
		
		/// <summary>
        /// Creates a new instance of Explorer
        /// </summary>
        ///<param name="progId">registered ProgID</param>
		public Explorer(string progId):base(progId)
		{
			
		}

		#endregion

		#region Static CoClass Methods
		#endregion

		#region Events

		/// <summary>
		/// SupportByVersion Outlook, 9,10,11,12,14,15,16
		/// </summary>
		private event Explorer_ActivateEventHandler _ActivateEvent;

		/// <summary>
		/// SupportByVersion Outlook 9 10 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff867298.aspx </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		public event Explorer_ActivateEventHandler ActivateEvent
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
		/// SupportByVersion Outlook, 9,10,11,12,14,15,16
		/// </summary>
		private event Explorer_FolderSwitchEventHandler _FolderSwitchEvent;

		/// <summary>
		/// SupportByVersion Outlook 9 10 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff865625.aspx </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		public event Explorer_FolderSwitchEventHandler FolderSwitchEvent
		{
			add
			{
				CreateEventBridge();
				_FolderSwitchEvent += value;
			}
			remove
			{
				_FolderSwitchEvent -= value;
			}
		}

		/// <summary>
		/// SupportByVersion Outlook, 9,10,11,12,14,15,16
		/// </summary>
		private event Explorer_BeforeFolderSwitchEventHandler _BeforeFolderSwitchEvent;

		/// <summary>
		/// SupportByVersion Outlook 9 10 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff868537.aspx </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		public event Explorer_BeforeFolderSwitchEventHandler BeforeFolderSwitchEvent
		{
			add
			{
				CreateEventBridge();
				_BeforeFolderSwitchEvent += value;
			}
			remove
			{
				_BeforeFolderSwitchEvent -= value;
			}
		}

		/// <summary>
		/// SupportByVersion Outlook, 9,10,11,12,14,15,16
		/// </summary>
		private event Explorer_ViewSwitchEventHandler _ViewSwitchEvent;

		/// <summary>
		/// SupportByVersion Outlook 9 10 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff868484.aspx </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		public event Explorer_ViewSwitchEventHandler ViewSwitchEvent
		{
			add
			{
				CreateEventBridge();
				_ViewSwitchEvent += value;
			}
			remove
			{
				_ViewSwitchEvent -= value;
			}
		}

		/// <summary>
		/// SupportByVersion Outlook, 9,10,11,12,14,15,16
		/// </summary>
		private event Explorer_BeforeViewSwitchEventHandler _BeforeViewSwitchEvent;

		/// <summary>
		/// SupportByVersion Outlook 9 10 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff865397.aspx </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		public event Explorer_BeforeViewSwitchEventHandler BeforeViewSwitchEvent
		{
			add
			{
				CreateEventBridge();
				_BeforeViewSwitchEvent += value;
			}
			remove
			{
				_BeforeViewSwitchEvent -= value;
			}
		}

		/// <summary>
		/// SupportByVersion Outlook, 9,10,11,12,14,15,16
		/// </summary>
		private event Explorer_DeactivateEventHandler _DeactivateEvent;

		/// <summary>
		/// SupportByVersion Outlook 9 10 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff866945.aspx </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		public event Explorer_DeactivateEventHandler DeactivateEvent
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
		/// SupportByVersion Outlook, 9,10,11,12,14,15,16
		/// </summary>
		private event Explorer_SelectionChangeEventHandler _SelectionChangeEvent;

		/// <summary>
		/// SupportByVersion Outlook 9 10 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff869813.aspx </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		public event Explorer_SelectionChangeEventHandler SelectionChangeEvent
		{
			add
			{
				CreateEventBridge();
				_SelectionChangeEvent += value;
			}
			remove
			{
				_SelectionChangeEvent -= value;
			}
		}

		/// <summary>
		/// SupportByVersion Outlook, 9,10,11,12,14,15,16
		/// </summary>
		private event Explorer_CloseEventHandler _CloseEvent;

		/// <summary>
		/// SupportByVersion Outlook 9 10 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff862184.aspx </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		public event Explorer_CloseEventHandler CloseEvent
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
		/// SupportByVersion Outlook, 10,11,12,14,15,16
		/// </summary>
		private event Explorer_BeforeMaximizeEventHandler _BeforeMaximizeEvent;

		/// <summary>
		/// SupportByVersion Outlook 10 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff864743.aspx </remarks>
		[SupportByVersion("Outlook", 10,11,12,14,15,16)]
		public event Explorer_BeforeMaximizeEventHandler BeforeMaximizeEvent
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
		/// SupportByVersion Outlook, 10,11,12,14,15,16
		/// </summary>
		private event Explorer_BeforeMinimizeEventHandler _BeforeMinimizeEvent;

		/// <summary>
		/// SupportByVersion Outlook 10 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff868043.aspx </remarks>
		[SupportByVersion("Outlook", 10,11,12,14,15,16)]
		public event Explorer_BeforeMinimizeEventHandler BeforeMinimizeEvent
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
		/// SupportByVersion Outlook, 10,11,12,14,15,16
		/// </summary>
		private event Explorer_BeforeMoveEventHandler _BeforeMoveEvent;

		/// <summary>
		/// SupportByVersion Outlook 10 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff868815.aspx </remarks>
		[SupportByVersion("Outlook", 10,11,12,14,15,16)]
		public event Explorer_BeforeMoveEventHandler BeforeMoveEvent
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
		/// SupportByVersion Outlook, 10,11,12,14,15,16
		/// </summary>
		private event Explorer_BeforeSizeEventHandler _BeforeSizeEvent;

		/// <summary>
		/// SupportByVersion Outlook 10 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff862995.aspx </remarks>
		[SupportByVersion("Outlook", 10,11,12,14,15,16)]
		public event Explorer_BeforeSizeEventHandler BeforeSizeEvent
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
		/// SupportByVersion Outlook, 10,11,12,14,15,16
		/// </summary>
		private event Explorer_BeforeItemCopyEventHandler _BeforeItemCopyEvent;

		/// <summary>
		/// SupportByVersion Outlook 10 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff860454.aspx </remarks>
		[SupportByVersion("Outlook", 10,11,12,14,15,16)]
		public event Explorer_BeforeItemCopyEventHandler BeforeItemCopyEvent
		{
			add
			{
				CreateEventBridge();
				_BeforeItemCopyEvent += value;
			}
			remove
			{
				_BeforeItemCopyEvent -= value;
			}
		}

		/// <summary>
		/// SupportByVersion Outlook, 10,11,12,14,15,16
		/// </summary>
		private event Explorer_BeforeItemCutEventHandler _BeforeItemCutEvent;

		/// <summary>
		/// SupportByVersion Outlook 10 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff867174.aspx </remarks>
		[SupportByVersion("Outlook", 10,11,12,14,15,16)]
		public event Explorer_BeforeItemCutEventHandler BeforeItemCutEvent
		{
			add
			{
				CreateEventBridge();
				_BeforeItemCutEvent += value;
			}
			remove
			{
				_BeforeItemCutEvent -= value;
			}
		}

		/// <summary>
		/// SupportByVersion Outlook, 10,11,12,14,15,16
		/// </summary>
		private event Explorer_BeforeItemPasteEventHandler _BeforeItemPasteEvent;

		/// <summary>
		/// SupportByVersion Outlook 10 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff868366.aspx </remarks>
		[SupportByVersion("Outlook", 10,11,12,14,15,16)]
		public event Explorer_BeforeItemPasteEventHandler BeforeItemPasteEvent
		{
			add
			{
				CreateEventBridge();
				_BeforeItemPasteEvent += value;
			}
			remove
			{
				_BeforeItemPasteEvent -= value;
			}
		}

		/// <summary>
		/// SupportByVersion Outlook, 14,15,16
		/// </summary>
		private event Explorer_AttachmentSelectionChangeEventHandler _AttachmentSelectionChangeEvent;

		/// <summary>
		/// SupportByVersion Outlook 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff867876.aspx </remarks>
		[SupportByVersion("Outlook", 14,15,16)]
		public event Explorer_AttachmentSelectionChangeEventHandler AttachmentSelectionChangeEvent
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

		/// <summary>
		/// SupportByVersion Outlook, 15, 16
		/// </summary>
		private event Explorer_InlineResponseEventHandler _InlineResponseEvent;

		/// <summary>
		/// SupportByVersion Outlook 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/jj229061.aspx </remarks>
		[SupportByVersion("Outlook", 15, 16)]
		public event Explorer_InlineResponseEventHandler InlineResponseEvent
		{
			add
			{
				CreateEventBridge();
				_InlineResponseEvent += value;
			}
			remove
			{
				_InlineResponseEvent -= value;
			}
		}

        /// <summary>
        /// SupportByVersion Outlook, 15, 16
        /// </summary>
        private event Explorer_InlineResponseCloseEventHandler _InlineResponseCloseEvent;

        /// <summary>
        /// SupportByVersion Outlook 15,16
        /// </summary>
        ///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/jj229061.aspx </remarks>
        [SupportByVersion("Outlook", 15, 16)]
        public event Explorer_InlineResponseCloseEventHandler InlineResponseCloseEvent
        {
            add
            {
                CreateEventBridge();
                _InlineResponseCloseEvent += value;
            }
            remove
            {
                _InlineResponseCloseEvent -= value;
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
				_activeSinkId = SinkHelper.GetConnectionPoint(this, ref _connectPoint, Events.ExplorerEvents_SinkHelper.Id, Events.ExplorerEvents_10_SinkHelper.Id);


			if(Events.ExplorerEvents_SinkHelper.Id.Equals(_activeSinkId, StringComparison.InvariantCultureIgnoreCase))
			{
				_explorerEvents_SinkHelper = new Events.ExplorerEvents_SinkHelper(this, _connectPoint);
				return;
			}

			if(Events.ExplorerEvents_10_SinkHelper.Id.Equals(_activeSinkId, StringComparison.InvariantCultureIgnoreCase))
			{
				_explorerEvents_10_SinkHelper = new Events.ExplorerEvents_10_SinkHelper(this, _connectPoint);
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
			if( null != _explorerEvents_SinkHelper)
			{
				_explorerEvents_SinkHelper.Dispose();
				_explorerEvents_SinkHelper = null;
			}
			if( null != _explorerEvents_10_SinkHelper)
			{
				_explorerEvents_10_SinkHelper.Dispose();
				_explorerEvents_10_SinkHelper = null;
			}

			_connectPoint = null;
		}
        
        #endregion

		#pragma warning restore
	}
}

