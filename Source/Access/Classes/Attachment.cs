﻿using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.AccessApi
{
	#region Delegates

	#pragma warning disable
	public delegate void Attachment_BeforeUpdateEventHandler(ref Int16 cancel);
	public delegate void Attachment_AfterUpdateEventHandler();
	public delegate void Attachment_EnterEventHandler();
	public delegate void Attachment_ExitEventHandler(ref Int16 cancel);
	public delegate void Attachment_DirtyEventHandler(ref Int16 cancel);
	public delegate void Attachment_ChangeEventHandler();
	public delegate void Attachment_GotFocusEventHandler();
	public delegate void Attachment_LostFocusEventHandler();
	public delegate void Attachment_ClickEventHandler();
	public delegate void Attachment_DblClickEventHandler(ref Int16 cancel);
	public delegate void Attachment_MouseDownEventHandler(ref Int16 button, ref Int16 shift, ref Single x, ref Single y);
	public delegate void Attachment_MouseMoveEventHandler(ref Int16 button, ref Int16 shift, ref Single x, ref Single y);
	public delegate void Attachment_MouseUpEventHandler(ref Int16 button, ref Int16 shift, ref Single x, ref Single y);
	public delegate void Attachment_KeyDownEventHandler(ref Int16 keyCode, ref Int16 shift);
	public delegate void Attachment_KeyPressEventHandler(ref Int16 keyAscii);
	public delegate void Attachment_KeyUpEventHandler(ref Int16 keyCode, ref Int16 shift);
	public delegate void Attachment_AttachmentCurrentEventHandler();
	#pragma warning restore

	#endregion

	/// <summary>
	/// CoClass Attachment 
	/// SupportByVersion Access, 12,14,15,16
	/// </summary>
	/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Access.Attachment"/> </remarks>
	[SupportByVersion("Access", 12,14,15,16)]
	[EntityType(EntityType.IsCoClass)]
    [EventSink(typeof(Events.DispAttachmentEvents_SinkHelper))]
    [ComEventInterface(typeof(Events.DispAttachmentEvents))]
    public class Attachment : _Attachment, IEventBinding
	{
		#pragma warning disable

		#region Fields
		
		private NetRuntimeSystem.Runtime.InteropServices.ComTypes.IConnectionPoint _connectPoint;
		private string _activeSinkId;
        private static Type _type;
        private Events.DispAttachmentEvents_SinkHelper _dispAttachmentEvents_SinkHelper;
	
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
                    _type = typeof(Attachment);
                return _type;
            }
        }
        
        #endregion
        		
		#region Construction

		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
		public Attachment(Core factory, ICOMObject parentObject, object comProxy) : base(factory, parentObject, comProxy)
		{
			
		}

        ///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
		public Attachment(ICOMObject parentObject, object comProxy) : base(parentObject, comProxy)
		{
			
		}

		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Attachment(Core factory, ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(factory, parentObject, comProxy, comProxyType)
		{
			
		}

		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Attachment(ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(parentObject, comProxy, comProxyType)
		{
			
		}
		
		///<param name="replacedObject">object to replaced. replacedObject are not usable after this action</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Attachment(ICOMObject replacedObject) : base(replacedObject)
		{
			
		}
		
		/// <summary>
        /// Creates a new instance of Attachment 
        /// </summary>		
		public Attachment():base("Access.Attachment")
		{
			
		}
		
		/// <summary>
        /// Creates a new instance of Attachment
        /// </summary>
        ///<param name="progId">registered ProgID</param>
		public Attachment(string progId):base(progId)
		{
			
		}

		#endregion

		#region Static CoClass Methods
		#endregion

		#region Events

		/// <summary>
		/// SupportByVersion Access, 12,14,15,16
		/// </summary>
		private event Attachment_BeforeUpdateEventHandler _BeforeUpdateEvent;

		/// <summary>
		/// SupportByVersion Access 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: <see href="https://docs.microsoft.com/en-us/office/vba/api/Access.Attachment.BeforeUpdate-event"/> </remarks>
		[SupportByVersion("Access", 12,14,15,16)]
		public event Attachment_BeforeUpdateEventHandler BeforeUpdateEvent
		{
			add
			{
				CreateEventBridge();
				_BeforeUpdateEvent += value;
			}
			remove
			{
				_BeforeUpdateEvent -= value;
			}
		}

		/// <summary>
		/// SupportByVersion Access, 12,14,15,16
		/// </summary>
		private event Attachment_AfterUpdateEventHandler _AfterUpdateEvent;

		/// <summary>
		/// SupportByVersion Access 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: <see href="https://docs.microsoft.com/en-us/office/vba/api/Access.Attachment.AfterUpdate-event"/> </remarks>
		[SupportByVersion("Access", 12,14,15,16)]
		public event Attachment_AfterUpdateEventHandler AfterUpdateEvent
		{
			add
			{
				CreateEventBridge();
				_AfterUpdateEvent += value;
			}
			remove
			{
				_AfterUpdateEvent -= value;
			}
		}

		/// <summary>
		/// SupportByVersion Access, 12,14,15,16
		/// </summary>
		private event Attachment_EnterEventHandler _EnterEvent;

		/// <summary>
		/// SupportByVersion Access 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: <see href="https://docs.microsoft.com/en-us/office/vba/api/Access.Attachment.Enter"/> </remarks>
		[SupportByVersion("Access", 12,14,15,16)]
		public event Attachment_EnterEventHandler EnterEvent
		{
			add
			{
				CreateEventBridge();
				_EnterEvent += value;
			}
			remove
			{
				_EnterEvent -= value;
			}
		}

		/// <summary>
		/// SupportByVersion Access, 12,14,15,16
		/// </summary>
		private event Attachment_ExitEventHandler _ExitEvent;

		/// <summary>
		/// SupportByVersion Access 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: <see href="https://docs.microsoft.com/en-us/office/vba/api/Access.Attachment.Exit"/> </remarks>
		[SupportByVersion("Access", 12,14,15,16)]
		public event Attachment_ExitEventHandler ExitEvent
		{
			add
			{
				CreateEventBridge();
				_ExitEvent += value;
			}
			remove
			{
				_ExitEvent -= value;
			}
		}

		/// <summary>
		/// SupportByVersion Access, 12,14,15,16
		/// </summary>
		private event Attachment_DirtyEventHandler _DirtyEvent;

		/// <summary>
		/// SupportByVersion Access 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: <see href="https://docs.microsoft.com/en-us/office/vba/api/Access.Attachment.Dirty"/> </remarks>
		[SupportByVersion("Access", 12,14,15,16)]
		public event Attachment_DirtyEventHandler DirtyEvent
		{
			add
			{
				CreateEventBridge();
				_DirtyEvent += value;
			}
			remove
			{
				_DirtyEvent -= value;
			}
		}

		/// <summary>
		/// SupportByVersion Access, 12,14,15,16
		/// </summary>
		private event Attachment_ChangeEventHandler _ChangeEvent;

		/// <summary>
		/// SupportByVersion Access 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: <see href="https://docs.microsoft.com/en-us/office/vba/api/Access.Attachment.Change"/> </remarks>
		[SupportByVersion("Access", 12,14,15,16)]
		public event Attachment_ChangeEventHandler ChangeEvent
		{
			add
			{
				CreateEventBridge();
				_ChangeEvent += value;
			}
			remove
			{
				_ChangeEvent -= value;
			}
		}

		/// <summary>
		/// SupportByVersion Access, 12,14,15,16
		/// </summary>
		private event Attachment_GotFocusEventHandler _GotFocusEvent;

		/// <summary>
		/// SupportByVersion Access 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: <see href="https://docs.microsoft.com/en-us/office/vba/api/Access.Attachment.GotFocus"/> </remarks>
		[SupportByVersion("Access", 12,14,15,16)]
		public event Attachment_GotFocusEventHandler GotFocusEvent
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
		/// SupportByVersion Access, 12,14,15,16
		/// </summary>
		private event Attachment_LostFocusEventHandler _LostFocusEvent;

		/// <summary>
		/// SupportByVersion Access 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: <see href="https://docs.microsoft.com/en-us/office/vba/api/Access.Attachment.LostFocus"/> </remarks>
		[SupportByVersion("Access", 12,14,15,16)]
		public event Attachment_LostFocusEventHandler LostFocusEvent
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

		/// <summary>
		/// SupportByVersion Access, 12,14,15,16
		/// </summary>
		private event Attachment_ClickEventHandler _ClickEvent;

		/// <summary>
		/// SupportByVersion Access 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: <see href="https://docs.microsoft.com/en-us/office/vba/api/Access.Attachment.Click"/> </remarks>
		[SupportByVersion("Access", 12,14,15,16)]
		public event Attachment_ClickEventHandler ClickEvent
		{
			add
			{
				CreateEventBridge();
				_ClickEvent += value;
			}
			remove
			{
				_ClickEvent -= value;
			}
		}

		/// <summary>
		/// SupportByVersion Access, 12,14,15,16
		/// </summary>
		private event Attachment_DblClickEventHandler _DblClickEvent;

		/// <summary>
		/// SupportByVersion Access 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: <see href="https://docs.microsoft.com/en-us/office/vba/api/Access.Attachment.DblClick"/> </remarks>
		[SupportByVersion("Access", 12,14,15,16)]
		public event Attachment_DblClickEventHandler DblClickEvent
		{
			add
			{
				CreateEventBridge();
				_DblClickEvent += value;
			}
			remove
			{
				_DblClickEvent -= value;
			}
		}

		/// <summary>
		/// SupportByVersion Access, 12,14,15,16
		/// </summary>
		private event Attachment_MouseDownEventHandler _MouseDownEvent;

		/// <summary>
		/// SupportByVersion Access 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: <see href="https://docs.microsoft.com/en-us/office/vba/api/Access.Attachment.MouseDown"/> </remarks>
		[SupportByVersion("Access", 12,14,15,16)]
		public event Attachment_MouseDownEventHandler MouseDownEvent
		{
			add
			{
				CreateEventBridge();
				_MouseDownEvent += value;
			}
			remove
			{
				_MouseDownEvent -= value;
			}
		}

		/// <summary>
		/// SupportByVersion Access, 12,14,15,16
		/// </summary>
		private event Attachment_MouseMoveEventHandler _MouseMoveEvent;

		/// <summary>
		/// SupportByVersion Access 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: <see href="https://docs.microsoft.com/en-us/office/vba/api/Access.Attachment.MouseMove"/> </remarks>
		[SupportByVersion("Access", 12,14,15,16)]
		public event Attachment_MouseMoveEventHandler MouseMoveEvent
		{
			add
			{
				CreateEventBridge();
				_MouseMoveEvent += value;
			}
			remove
			{
				_MouseMoveEvent -= value;
			}
		}

		/// <summary>
		/// SupportByVersion Access, 12,14,15,16
		/// </summary>
		private event Attachment_MouseUpEventHandler _MouseUpEvent;

		/// <summary>
		/// SupportByVersion Access 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: <see href="https://docs.microsoft.com/en-us/office/vba/api/Access.Attachment.MouseUp"/> </remarks>
		[SupportByVersion("Access", 12,14,15,16)]
		public event Attachment_MouseUpEventHandler MouseUpEvent
		{
			add
			{
				CreateEventBridge();
				_MouseUpEvent += value;
			}
			remove
			{
				_MouseUpEvent -= value;
			}
		}

		/// <summary>
		/// SupportByVersion Access, 12,14,15,16
		/// </summary>
		private event Attachment_KeyDownEventHandler _KeyDownEvent;

		/// <summary>
		/// SupportByVersion Access 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: <see href="https://docs.microsoft.com/en-us/office/vba/api/Access.Attachment.KeyDown"/> </remarks>
		[SupportByVersion("Access", 12,14,15,16)]
		public event Attachment_KeyDownEventHandler KeyDownEvent
		{
			add
			{
				CreateEventBridge();
				_KeyDownEvent += value;
			}
			remove
			{
				_KeyDownEvent -= value;
			}
		}

		/// <summary>
		/// SupportByVersion Access, 12,14,15,16
		/// </summary>
		private event Attachment_KeyPressEventHandler _KeyPressEvent;

		/// <summary>
		/// SupportByVersion Access 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: <see href="https://docs.microsoft.com/en-us/office/vba/api/Access.Attachment.KeyPress"/> </remarks>
		[SupportByVersion("Access", 12,14,15,16)]
		public event Attachment_KeyPressEventHandler KeyPressEvent
		{
			add
			{
				CreateEventBridge();
				_KeyPressEvent += value;
			}
			remove
			{
				_KeyPressEvent -= value;
			}
		}

		/// <summary>
		/// SupportByVersion Access, 12,14,15,16
		/// </summary>
		private event Attachment_KeyUpEventHandler _KeyUpEvent;

		/// <summary>
		/// SupportByVersion Access 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: <see href="https://docs.microsoft.com/en-us/office/vba/api/Access.Attachment.KeyUp"/> </remarks>
		[SupportByVersion("Access", 12,14,15,16)]
		public event Attachment_KeyUpEventHandler KeyUpEvent
		{
			add
			{
				CreateEventBridge();
				_KeyUpEvent += value;
			}
			remove
			{
				_KeyUpEvent -= value;
			}
		}

		/// <summary>
		/// SupportByVersion Access, 12,14,15,16
		/// </summary>
		private event Attachment_AttachmentCurrentEventHandler _AttachmentCurrentEvent;

		/// <summary>
		/// SupportByVersion Access 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: <see href="https://docs.microsoft.com/en-us/office/vba/api/Access.Attachment.AttachmentCurrent"/> </remarks>
		[SupportByVersion("Access", 12,14,15,16)]
		public event Attachment_AttachmentCurrentEventHandler AttachmentCurrentEvent
		{
			add
			{
				CreateEventBridge();
				_AttachmentCurrentEvent += value;
			}
			remove
			{
				_AttachmentCurrentEvent -= value;
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
				_activeSinkId = SinkHelper.GetConnectionPoint(this, ref _connectPoint, Events.DispAttachmentEvents_SinkHelper.Id);


			if(Events.DispAttachmentEvents_SinkHelper.Id.Equals(_activeSinkId, StringComparison.InvariantCultureIgnoreCase))
			{
				_dispAttachmentEvents_SinkHelper = new Events.DispAttachmentEvents_SinkHelper(this, _connectPoint);
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
			if( null != _dispAttachmentEvents_SinkHelper)
			{
				_dispAttachmentEvents_SinkHelper.Dispose();
				_dispAttachmentEvents_SinkHelper = null;
			}

			_connectPoint = null;
		}
        
        #endregion

		#pragma warning restore
	}
}

