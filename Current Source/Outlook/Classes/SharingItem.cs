using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using LateBindingApi.Core;
namespace NetOffice.OutlookApi
{

	#region Delegates

	#pragma warning disable
	public delegate void SharingItem_OpenEventHandler(ref bool Cancel);
	public delegate void SharingItem_CustomActionEventHandler(COMObject Action, COMObject Response, ref bool Cancel);
	public delegate void SharingItem_CustomPropertyChangeEventHandler(string Name);
	public delegate void SharingItem_ForwardEventHandler(COMObject Forward, ref bool Cancel);
	public delegate void SharingItem_CloseEventHandler(ref bool Cancel);
	public delegate void SharingItem_PropertyChangeEventHandler(string Name);
	public delegate void SharingItem_ReadEventHandler();
	public delegate void SharingItem_ReplyEventHandler(COMObject Response, ref bool Cancel);
	public delegate void SharingItem_ReplyAllEventHandler(COMObject Response, ref bool Cancel);
	public delegate void SharingItem_SendEventHandler(ref bool Cancel);
	public delegate void SharingItem_WriteEventHandler(ref bool Cancel);
	public delegate void SharingItem_BeforeCheckNamesEventHandler(ref bool Cancel);
	public delegate void SharingItem_AttachmentAddEventHandler(NetOffice.OutlookApi.Attachment Attachment);
	public delegate void SharingItem_AttachmentReadEventHandler(NetOffice.OutlookApi.Attachment Attachment);
	public delegate void SharingItem_BeforeAttachmentSaveEventHandler(NetOffice.OutlookApi.Attachment Attachment, ref bool Cancel);
	public delegate void SharingItem_BeforeDeleteEventHandler(COMObject Item, ref bool Cancel);
	public delegate void SharingItem_AttachmentRemoveEventHandler(NetOffice.OutlookApi.Attachment Attachment);
	public delegate void SharingItem_BeforeAttachmentAddEventHandler(NetOffice.OutlookApi.Attachment Attachment, ref bool Cancel);
	public delegate void SharingItem_BeforeAttachmentPreviewEventHandler(NetOffice.OutlookApi.Attachment Attachment, ref bool Cancel);
	public delegate void SharingItem_BeforeAttachmentReadEventHandler(NetOffice.OutlookApi.Attachment Attachment, ref bool Cancel);
	public delegate void SharingItem_BeforeAttachmentWriteToTempFileEventHandler(NetOffice.OutlookApi.Attachment Attachment, ref bool Cancel);
	public delegate void SharingItem_UnloadEventHandler();
	public delegate void SharingItem_BeforeAutoSaveEventHandler(ref bool Cancel);
	public delegate void SharingItem_BeforeReadEventHandler();
	public delegate void SharingItem_AfterWriteEventHandler();
	#pragma warning restore

	#endregion

	///<summary>
	/// CoClass SharingItem 
	/// SupportByVersion Outlook, 12,14
	///</summary>
	[SupportByVersionAttribute("Outlook", 12,14)]
	[EntityTypeAttribute(EntityType.IsCoClass)]
	public class SharingItem : _SharingItem,IEventBinding
	{
		#pragma warning disable
		#region Fields
		
		private NetRuntimeSystem.Runtime.InteropServices.ComTypes.IConnectionPoint _connectPoint;
		private string _activeSinkId;
		private NetRuntimeSystem.Type _thisType;
		ItemEvents_10_SinkHelper _itemEvents_10_SinkHelper;
	
		#endregion

		#region Type Information

        private static Type _type;
		
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        public static Type LateBindingApiWrapperType
        {
            get
            {
                if (null == _type)
                    _type = typeof(SharingItem);
                    
                return _type;
            }
        }
        
        #endregion
        		
		#region Construction

        /// <param name="parentObject">object there has created the proxy</param>
        /// <param name="comProxy">inner wrapped COM proxy</param>
		public SharingItem(COMObject parentObject, object comProxy) : base(parentObject, comProxy)
		{
			
		}

		/// <param name="parentObject">object there has created the proxy</param>
        /// <param name="comProxy">inner wrapped COM proxy</param>
        /// <param name="comProxyType">Type of inner wrapped COM proxy"</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public SharingItem(COMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(parentObject, comProxy, comProxyType)
		{
			
		}
		
		/// <param name="replacedObject">object to replaced. replacedObject are not usable after this action</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public SharingItem(COMObject replacedObject) : base(replacedObject)
		{
			
		}
		
		/// <summary>
        /// creates a new instance of SharingItem 
        /// </summary>		
		public SharingItem():base("Outlook.SharingItem")
		{
			
		}
		
		/// <summary>
        /// creates a new instance of SharingItem
        /// </summary>
        /// <param name="progId">registered ProgID</param>
		public SharingItem(string progId):base(progId)
		{
			
		}

		#endregion

		#region Events

		/// <summary>
		/// SupportByVersion Outlook, 10,11,12,14
		/// </summary>
		private event SharingItem_OpenEventHandler _OpenEvent;

		/// <summary>
		/// SupportByVersion Outlook 10 11 12 14
		/// </summary>
		[SupportByVersion("Outlook", 10,11,12,14)]
		public event SharingItem_OpenEventHandler OpenEvent
		{
			add
			{
				CreateEventBridge();
				_OpenEvent += value;
			}
			remove
			{
				_OpenEvent -= value;
			}
		}

		/// <summary>
		/// SupportByVersion Outlook, 10,11,12,14
		/// </summary>
		private event SharingItem_CustomActionEventHandler _CustomActionEvent;

		/// <summary>
		/// SupportByVersion Outlook 10 11 12 14
		/// </summary>
		[SupportByVersion("Outlook", 10,11,12,14)]
		public event SharingItem_CustomActionEventHandler CustomActionEvent
		{
			add
			{
				CreateEventBridge();
				_CustomActionEvent += value;
			}
			remove
			{
				_CustomActionEvent -= value;
			}
		}

		/// <summary>
		/// SupportByVersion Outlook, 10,11,12,14
		/// </summary>
		private event SharingItem_CustomPropertyChangeEventHandler _CustomPropertyChangeEvent;

		/// <summary>
		/// SupportByVersion Outlook 10 11 12 14
		/// </summary>
		[SupportByVersion("Outlook", 10,11,12,14)]
		public event SharingItem_CustomPropertyChangeEventHandler CustomPropertyChangeEvent
		{
			add
			{
				CreateEventBridge();
				_CustomPropertyChangeEvent += value;
			}
			remove
			{
				_CustomPropertyChangeEvent -= value;
			}
		}

		/// <summary>
		/// SupportByVersion Outlook, 10,11,12,14
		/// </summary>
		private event SharingItem_ForwardEventHandler _ForwardEvent;

		/// <summary>
		/// SupportByVersion Outlook 10 11 12 14
		/// </summary>
		[SupportByVersion("Outlook", 10,11,12,14)]
		public event SharingItem_ForwardEventHandler ForwardEvent
		{
			add
			{
				CreateEventBridge();
				_ForwardEvent += value;
			}
			remove
			{
				_ForwardEvent -= value;
			}
		}

		/// <summary>
		/// SupportByVersion Outlook, 10,11,12,14
		/// </summary>
		private event SharingItem_CloseEventHandler _CloseEvent;

		/// <summary>
		/// SupportByVersion Outlook 10 11 12 14
		/// </summary>
		[SupportByVersion("Outlook", 10,11,12,14)]
		public event SharingItem_CloseEventHandler CloseEvent
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
		/// SupportByVersion Outlook, 10,11,12,14
		/// </summary>
		private event SharingItem_PropertyChangeEventHandler _PropertyChangeEvent;

		/// <summary>
		/// SupportByVersion Outlook 10 11 12 14
		/// </summary>
		[SupportByVersion("Outlook", 10,11,12,14)]
		public event SharingItem_PropertyChangeEventHandler PropertyChangeEvent
		{
			add
			{
				CreateEventBridge();
				_PropertyChangeEvent += value;
			}
			remove
			{
				_PropertyChangeEvent -= value;
			}
		}

		/// <summary>
		/// SupportByVersion Outlook, 10,11,12,14
		/// </summary>
		private event SharingItem_ReadEventHandler _ReadEvent;

		/// <summary>
		/// SupportByVersion Outlook 10 11 12 14
		/// </summary>
		[SupportByVersion("Outlook", 10,11,12,14)]
		public event SharingItem_ReadEventHandler ReadEvent
		{
			add
			{
				CreateEventBridge();
				_ReadEvent += value;
			}
			remove
			{
				_ReadEvent -= value;
			}
		}

		/// <summary>
		/// SupportByVersion Outlook, 10,11,12,14
		/// </summary>
		private event SharingItem_ReplyEventHandler _ReplyEvent;

		/// <summary>
		/// SupportByVersion Outlook 10 11 12 14
		/// </summary>
		[SupportByVersion("Outlook", 10,11,12,14)]
		public event SharingItem_ReplyEventHandler ReplyEvent
		{
			add
			{
				CreateEventBridge();
				_ReplyEvent += value;
			}
			remove
			{
				_ReplyEvent -= value;
			}
		}

		/// <summary>
		/// SupportByVersion Outlook, 10,11,12,14
		/// </summary>
		private event SharingItem_ReplyAllEventHandler _ReplyAllEvent;

		/// <summary>
		/// SupportByVersion Outlook 10 11 12 14
		/// </summary>
		[SupportByVersion("Outlook", 10,11,12,14)]
		public event SharingItem_ReplyAllEventHandler ReplyAllEvent
		{
			add
			{
				CreateEventBridge();
				_ReplyAllEvent += value;
			}
			remove
			{
				_ReplyAllEvent -= value;
			}
		}

		/// <summary>
		/// SupportByVersion Outlook, 10,11,12,14
		/// </summary>
		private event SharingItem_SendEventHandler _SendEvent;

		/// <summary>
		/// SupportByVersion Outlook 10 11 12 14
		/// </summary>
		[SupportByVersion("Outlook", 10,11,12,14)]
		public event SharingItem_SendEventHandler SendEvent
		{
			add
			{
				CreateEventBridge();
				_SendEvent += value;
			}
			remove
			{
				_SendEvent -= value;
			}
		}

		/// <summary>
		/// SupportByVersion Outlook, 10,11,12,14
		/// </summary>
		private event SharingItem_WriteEventHandler _WriteEvent;

		/// <summary>
		/// SupportByVersion Outlook 10 11 12 14
		/// </summary>
		[SupportByVersion("Outlook", 10,11,12,14)]
		public event SharingItem_WriteEventHandler WriteEvent
		{
			add
			{
				CreateEventBridge();
				_WriteEvent += value;
			}
			remove
			{
				_WriteEvent -= value;
			}
		}

		/// <summary>
		/// SupportByVersion Outlook, 10,11,12,14
		/// </summary>
		private event SharingItem_BeforeCheckNamesEventHandler _BeforeCheckNamesEvent;

		/// <summary>
		/// SupportByVersion Outlook 10 11 12 14
		/// </summary>
		[SupportByVersion("Outlook", 10,11,12,14)]
		public event SharingItem_BeforeCheckNamesEventHandler BeforeCheckNamesEvent
		{
			add
			{
				CreateEventBridge();
				_BeforeCheckNamesEvent += value;
			}
			remove
			{
				_BeforeCheckNamesEvent -= value;
			}
		}

		/// <summary>
		/// SupportByVersion Outlook, 10,11,12,14
		/// </summary>
		private event SharingItem_AttachmentAddEventHandler _AttachmentAddEvent;

		/// <summary>
		/// SupportByVersion Outlook 10 11 12 14
		/// </summary>
		[SupportByVersion("Outlook", 10,11,12,14)]
		public event SharingItem_AttachmentAddEventHandler AttachmentAddEvent
		{
			add
			{
				CreateEventBridge();
				_AttachmentAddEvent += value;
			}
			remove
			{
				_AttachmentAddEvent -= value;
			}
		}

		/// <summary>
		/// SupportByVersion Outlook, 10,11,12,14
		/// </summary>
		private event SharingItem_AttachmentReadEventHandler _AttachmentReadEvent;

		/// <summary>
		/// SupportByVersion Outlook 10 11 12 14
		/// </summary>
		[SupportByVersion("Outlook", 10,11,12,14)]
		public event SharingItem_AttachmentReadEventHandler AttachmentReadEvent
		{
			add
			{
				CreateEventBridge();
				_AttachmentReadEvent += value;
			}
			remove
			{
				_AttachmentReadEvent -= value;
			}
		}

		/// <summary>
		/// SupportByVersion Outlook, 10,11,12,14
		/// </summary>
		private event SharingItem_BeforeAttachmentSaveEventHandler _BeforeAttachmentSaveEvent;

		/// <summary>
		/// SupportByVersion Outlook 10 11 12 14
		/// </summary>
		[SupportByVersion("Outlook", 10,11,12,14)]
		public event SharingItem_BeforeAttachmentSaveEventHandler BeforeAttachmentSaveEvent
		{
			add
			{
				CreateEventBridge();
				_BeforeAttachmentSaveEvent += value;
			}
			remove
			{
				_BeforeAttachmentSaveEvent -= value;
			}
		}

		/// <summary>
		/// SupportByVersion Outlook, 10,11,12,14
		/// </summary>
		private event SharingItem_BeforeDeleteEventHandler _BeforeDeleteEvent;

		/// <summary>
		/// SupportByVersion Outlook 10 11 12 14
		/// </summary>
		[SupportByVersion("Outlook", 10,11,12,14)]
		public event SharingItem_BeforeDeleteEventHandler BeforeDeleteEvent
		{
			add
			{
				CreateEventBridge();
				_BeforeDeleteEvent += value;
			}
			remove
			{
				_BeforeDeleteEvent -= value;
			}
		}

		/// <summary>
		/// SupportByVersion Outlook, 12,14
		/// </summary>
		private event SharingItem_AttachmentRemoveEventHandler _AttachmentRemoveEvent;

		/// <summary>
		/// SupportByVersion Outlook 12 14
		/// </summary>
		[SupportByVersion("Outlook", 12,14)]
		public event SharingItem_AttachmentRemoveEventHandler AttachmentRemoveEvent
		{
			add
			{
				CreateEventBridge();
				_AttachmentRemoveEvent += value;
			}
			remove
			{
				_AttachmentRemoveEvent -= value;
			}
		}

		/// <summary>
		/// SupportByVersion Outlook, 12,14
		/// </summary>
		private event SharingItem_BeforeAttachmentAddEventHandler _BeforeAttachmentAddEvent;

		/// <summary>
		/// SupportByVersion Outlook 12 14
		/// </summary>
		[SupportByVersion("Outlook", 12,14)]
		public event SharingItem_BeforeAttachmentAddEventHandler BeforeAttachmentAddEvent
		{
			add
			{
				CreateEventBridge();
				_BeforeAttachmentAddEvent += value;
			}
			remove
			{
				_BeforeAttachmentAddEvent -= value;
			}
		}

		/// <summary>
		/// SupportByVersion Outlook, 12,14
		/// </summary>
		private event SharingItem_BeforeAttachmentPreviewEventHandler _BeforeAttachmentPreviewEvent;

		/// <summary>
		/// SupportByVersion Outlook 12 14
		/// </summary>
		[SupportByVersion("Outlook", 12,14)]
		public event SharingItem_BeforeAttachmentPreviewEventHandler BeforeAttachmentPreviewEvent
		{
			add
			{
				CreateEventBridge();
				_BeforeAttachmentPreviewEvent += value;
			}
			remove
			{
				_BeforeAttachmentPreviewEvent -= value;
			}
		}

		/// <summary>
		/// SupportByVersion Outlook, 12,14
		/// </summary>
		private event SharingItem_BeforeAttachmentReadEventHandler _BeforeAttachmentReadEvent;

		/// <summary>
		/// SupportByVersion Outlook 12 14
		/// </summary>
		[SupportByVersion("Outlook", 12,14)]
		public event SharingItem_BeforeAttachmentReadEventHandler BeforeAttachmentReadEvent
		{
			add
			{
				CreateEventBridge();
				_BeforeAttachmentReadEvent += value;
			}
			remove
			{
				_BeforeAttachmentReadEvent -= value;
			}
		}

		/// <summary>
		/// SupportByVersion Outlook, 12,14
		/// </summary>
		private event SharingItem_BeforeAttachmentWriteToTempFileEventHandler _BeforeAttachmentWriteToTempFileEvent;

		/// <summary>
		/// SupportByVersion Outlook 12 14
		/// </summary>
		[SupportByVersion("Outlook", 12,14)]
		public event SharingItem_BeforeAttachmentWriteToTempFileEventHandler BeforeAttachmentWriteToTempFileEvent
		{
			add
			{
				CreateEventBridge();
				_BeforeAttachmentWriteToTempFileEvent += value;
			}
			remove
			{
				_BeforeAttachmentWriteToTempFileEvent -= value;
			}
		}

		/// <summary>
		/// SupportByVersion Outlook, 12,14
		/// </summary>
		private event SharingItem_UnloadEventHandler _UnloadEvent;

		/// <summary>
		/// SupportByVersion Outlook 12 14
		/// </summary>
		[SupportByVersion("Outlook", 12,14)]
		public event SharingItem_UnloadEventHandler UnloadEvent
		{
			add
			{
				CreateEventBridge();
				_UnloadEvent += value;
			}
			remove
			{
				_UnloadEvent -= value;
			}
		}

		/// <summary>
		/// SupportByVersion Outlook, 12,14
		/// </summary>
		private event SharingItem_BeforeAutoSaveEventHandler _BeforeAutoSaveEvent;

		/// <summary>
		/// SupportByVersion Outlook 12 14
		/// </summary>
		[SupportByVersion("Outlook", 12,14)]
		public event SharingItem_BeforeAutoSaveEventHandler BeforeAutoSaveEvent
		{
			add
			{
				CreateEventBridge();
				_BeforeAutoSaveEvent += value;
			}
			remove
			{
				_BeforeAutoSaveEvent -= value;
			}
		}

		/// <summary>
		/// SupportByVersion Outlook, 14
		/// </summary>
		private event SharingItem_BeforeReadEventHandler _BeforeReadEvent;

		/// <summary>
		/// SupportByVersion Outlook 14
		/// </summary>
		[SupportByVersion("Outlook", 14)]
		public event SharingItem_BeforeReadEventHandler BeforeReadEvent
		{
			add
			{
				CreateEventBridge();
				_BeforeReadEvent += value;
			}
			remove
			{
				_BeforeReadEvent -= value;
			}
		}

		/// <summary>
		/// SupportByVersion Outlook, 14
		/// </summary>
		private event SharingItem_AfterWriteEventHandler _AfterWriteEvent;

		/// <summary>
		/// SupportByVersion Outlook 14
		/// </summary>
		[SupportByVersion("Outlook", 14)]
		public event SharingItem_AfterWriteEventHandler AfterWriteEvent
		{
			add
			{
				CreateEventBridge();
				_AfterWriteEvent += value;
			}
			remove
			{
				_AfterWriteEvent -= value;
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
			if(false == LateBindingApi.Core.Settings.EnableEvents)
				return;
	
			if (null != _connectPoint)
				return;
	
            if (null == _activeSinkId)
				_activeSinkId = SinkHelper.GetConnectionPoint(this, ref _connectPoint, ItemEvents_10_SinkHelper.Id);


			if(ItemEvents_10_SinkHelper.Id.Equals(_activeSinkId, StringComparison.InvariantCultureIgnoreCase))
			{
				_itemEvents_10_SinkHelper = new ItemEvents_10_SinkHelper(this, _connectPoint);
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
                        DebugConsole.WriteException(exception);
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
			if( null != _itemEvents_10_SinkHelper)
			{
				_itemEvents_10_SinkHelper.Dispose();
				_itemEvents_10_SinkHelper = null;
			}

			_connectPoint = null;
		}
        
        #endregion

		#pragma warning restore
	}
}