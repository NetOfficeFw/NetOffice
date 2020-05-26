﻿using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.AccessApi
{
	#region Delegates

	#pragma warning disable
	public delegate void FormOldV10_LoadEventHandler();
	public delegate void FormOldV10_CurrentEventHandler();
	public delegate void FormOldV10_BeforeInsertEventHandler(ref Int16 cancel);
	public delegate void FormOldV10_AfterInsertEventHandler();
	public delegate void FormOldV10_BeforeUpdateEventHandler(ref Int16 cancel);
	public delegate void FormOldV10_AfterUpdateEventHandler();
	public delegate void FormOldV10_DeleteEventHandler(ref Int16 cancel);
	public delegate void FormOldV10_BeforeDelConfirmEventHandler(ref Int16 cancel, ref Int16 response);
	public delegate void FormOldV10_AfterDelConfirmEventHandler(ref Int16 status);
	public delegate void FormOldV10_OpenEventHandler(ref Int16 cancel);
	public delegate void FormOldV10_ResizeEventHandler();
	public delegate void FormOldV10_UnloadEventHandler(ref Int16 cancel);
	public delegate void FormOldV10_CloseEventHandler();
	public delegate void FormOldV10_ActivateEventHandler();
	public delegate void FormOldV10_DeactivateEventHandler();
	public delegate void FormOldV10_GotFocusEventHandler();
	public delegate void FormOldV10_LostFocusEventHandler();
	public delegate void FormOldV10_ClickEventHandler();
	public delegate void FormOldV10_DblClickEventHandler(ref Int16 cancel);
	public delegate void FormOldV10_MouseDownEventHandler(ref Int16 button, ref Int16 shift, ref Single x, ref Single y);
	public delegate void FormOldV10_MouseMoveEventHandler(ref Int16 button, ref Int16 shift, ref Single x, ref Single y);
	public delegate void FormOldV10_MouseUpEventHandler(ref Int16 button, ref Int16 shift, ref Single x, ref Single y);
	public delegate void FormOldV10_KeyDownEventHandler(ref Int16 keyCode, ref Int16 shift);
	public delegate void FormOldV10_KeyPressEventHandler(ref Int16 keyAscii);
	public delegate void FormOldV10_KeyUpEventHandler(ref Int16 keyCode, ref Int16 shift);
	public delegate void FormOldV10_ErrorEventHandler(ref Int16 dataErr, ref Int16 response);
	public delegate void FormOldV10_TimerEventHandler();
	public delegate void FormOldV10_FilterEventHandler(ref Int16 cancel, ref Int16 filterType);
	public delegate void FormOldV10_ApplyFilterEventHandler(ref Int16 cancel, ref Int16 applyType);
	public delegate void FormOldV10_DirtyEventHandler(ref Int16 cancel);
	public delegate void FormOldV10_UndoEventHandler(ref Int16 cancel);
	public delegate void FormOldV10_RecordExitEventHandler(ref Int16 cancel);
	public delegate void FormOldV10_BeginBatchEditEventHandler(ref Int16 cancel);
	public delegate void FormOldV10_UndoBatchEditEventHandler(ref Int16 cancel);
	public delegate void FormOldV10_BeforeBeginTransactionEventHandler(ref Int16 cancel, ref NetOffice.ADODBApi.Connection connection);
	public delegate void FormOldV10_AfterBeginTransactionEventHandler(ref NetOffice.ADODBApi.Connection connection);
	public delegate void FormOldV10_BeforeCommitTransactionEventHandler(ref Int16 cancel, ref NetOffice.ADODBApi.Connection connection);
	public delegate void FormOldV10_AfterCommitTransactionEventHandler(ref NetOffice.ADODBApi.Connection connection);
	public delegate void FormOldV10_RollbackTransactionEventHandler(ref NetOffice.ADODBApi.Connection connection);
	public delegate void FormOldV10_OnConnectEventHandler();
	public delegate void FormOldV10_OnDisconnectEventHandler();
	public delegate void FormOldV10_PivotTableChangeEventHandler(Int32 reason);
	public delegate void FormOldV10_QueryEventHandler();
	public delegate void FormOldV10_BeforeQueryEventHandler();
	public delegate void FormOldV10_SelectionChangeEventHandler();
	public delegate void FormOldV10_CommandBeforeExecuteEventHandler(object command, ICOMObject cancel);
	public delegate void FormOldV10_CommandCheckedEventHandler(object command, ICOMObject _checked);
	public delegate void FormOldV10_CommandEnabledEventHandler(object command, ICOMObject enabled);
	public delegate void FormOldV10_CommandExecuteEventHandler(object command);
	public delegate void FormOldV10_DataSetChangeEventHandler();
	public delegate void FormOldV10_BeforeScreenTipEventHandler(ICOMObject screenTipText, ICOMObject sourceObject);
	public delegate void FormOldV10_BeforeRenderEventHandler(ICOMObject drawObject, ICOMObject chartObject, ICOMObject cancel);
	public delegate void FormOldV10_AfterRenderEventHandler(ICOMObject drawObject, ICOMObject chartObject);
	public delegate void FormOldV10_AfterFinalRenderEventHandler(ICOMObject drawObject);
	public delegate void FormOldV10_AfterLayoutEventHandler(ICOMObject drawObject);
	public delegate void FormOldV10_MouseWheelEventHandler(bool page, Int32 count);
	public delegate void FormOldV10_ViewChangeEventHandler(Int32 reason);
	public delegate void FormOldV10_DataChangeEventHandler(Int32 reason);
	#pragma warning restore

	#endregion

	/// <summary>
	/// CoClass FormOldV10 
	/// SupportByVersion Access, 12,14,15,16
	/// </summary>
	[SupportByVersion("Access", 12,14,15,16)]
	[EntityType(EntityType.IsCoClass)]
    [EventSink(typeof(Events._FormEvents_SinkHelper))]
    [ComEventInterface(typeof(Events._FormEvents))]
    public class FormOldV10 : _Form2, IEventBinding
	{
		#pragma warning disable

		#region Fields
		
		private NetRuntimeSystem.Runtime.InteropServices.ComTypes.IConnectionPoint _connectPoint;
		private string _activeSinkId;
        private static Type _type;
        private Events._FormEvents_SinkHelper __FormEvents_SinkHelper;
	
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
                    _type = typeof(FormOldV10);
                return _type;
            }
        }
        
        #endregion
        		
		#region Construction

		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
		public FormOldV10(Core factory, ICOMObject parentObject, object comProxy) : base(factory, parentObject, comProxy)
		{
			
		}

        ///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
		public FormOldV10(ICOMObject parentObject, object comProxy) : base(parentObject, comProxy)
		{
			
		}

		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public FormOldV10(Core factory, ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(factory, parentObject, comProxy, comProxyType)
		{
			
		}

		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public FormOldV10(ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(parentObject, comProxy, comProxyType)
		{
			
		}
		
		///<param name="replacedObject">object to replaced. replacedObject are not usable after this action</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public FormOldV10(ICOMObject replacedObject) : base(replacedObject)
		{
			
		}
		
		/// <summary>
        /// Creates a new instance of FormOldV10 
        /// </summary>		
		public FormOldV10():base("Access.FormOldV10")
		{
			
		}
		
		/// <summary>
        /// Creates a new instance of FormOldV10
        /// </summary>
        ///<param name="progId">registered ProgID</param>
		public FormOldV10(string progId):base(progId)
		{
			
		}

		#endregion

		#region Static CoClass Methods
		#endregion

		#region Events

		/// <summary>
		/// SupportByVersion Access, 9,10,11,12,14,15,16
		/// </summary>
		private event FormOldV10_LoadEventHandler _LoadEvent;

		/// <summary>
		/// SupportByVersion Access 9 10 11 12 14 15,16
		/// </summary>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public event FormOldV10_LoadEventHandler LoadEvent
		{
			add
			{
				CreateEventBridge();
				_LoadEvent += value;
			}
			remove
			{
				_LoadEvent -= value;
			}
		}

		/// <summary>
		/// SupportByVersion Access, 9,10,11,12,14,15,16
		/// </summary>
		private event FormOldV10_CurrentEventHandler _CurrentEvent;

		/// <summary>
		/// SupportByVersion Access 9 10 11 12 14 15,16
		/// </summary>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public event FormOldV10_CurrentEventHandler CurrentEvent
		{
			add
			{
				CreateEventBridge();
				_CurrentEvent += value;
			}
			remove
			{
				_CurrentEvent -= value;
			}
		}

		/// <summary>
		/// SupportByVersion Access, 9,10,11,12,14,15,16
		/// </summary>
		private event FormOldV10_BeforeInsertEventHandler _BeforeInsertEvent;

		/// <summary>
		/// SupportByVersion Access 9 10 11 12 14 15,16
		/// </summary>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public event FormOldV10_BeforeInsertEventHandler BeforeInsertEvent
		{
			add
			{
				CreateEventBridge();
				_BeforeInsertEvent += value;
			}
			remove
			{
				_BeforeInsertEvent -= value;
			}
		}

		/// <summary>
		/// SupportByVersion Access, 9,10,11,12,14,15,16
		/// </summary>
		private event FormOldV10_AfterInsertEventHandler _AfterInsertEvent;

		/// <summary>
		/// SupportByVersion Access 9 10 11 12 14 15,16
		/// </summary>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public event FormOldV10_AfterInsertEventHandler AfterInsertEvent
		{
			add
			{
				CreateEventBridge();
				_AfterInsertEvent += value;
			}
			remove
			{
				_AfterInsertEvent -= value;
			}
		}

		/// <summary>
		/// SupportByVersion Access, 9,10,11,12,14,15,16
		/// </summary>
		private event FormOldV10_BeforeUpdateEventHandler _BeforeUpdateEvent;

		/// <summary>
		/// SupportByVersion Access 9 10 11 12 14 15,16
		/// </summary>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public event FormOldV10_BeforeUpdateEventHandler BeforeUpdateEvent
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
		/// SupportByVersion Access, 9,10,11,12,14,15,16
		/// </summary>
		private event FormOldV10_AfterUpdateEventHandler _AfterUpdateEvent;

		/// <summary>
		/// SupportByVersion Access 9 10 11 12 14 15,16
		/// </summary>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public event FormOldV10_AfterUpdateEventHandler AfterUpdateEvent
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
		/// SupportByVersion Access, 9,10,11,12,14,15,16
		/// </summary>
		private event FormOldV10_DeleteEventHandler _DeleteEvent;

		/// <summary>
		/// SupportByVersion Access 9 10 11 12 14 15,16
		/// </summary>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public event FormOldV10_DeleteEventHandler DeleteEvent
		{
			add
			{
				CreateEventBridge();
				_DeleteEvent += value;
			}
			remove
			{
				_DeleteEvent -= value;
			}
		}

		/// <summary>
		/// SupportByVersion Access, 9,10,11,12,14,15,16
		/// </summary>
		private event FormOldV10_BeforeDelConfirmEventHandler _BeforeDelConfirmEvent;

		/// <summary>
		/// SupportByVersion Access 9 10 11 12 14 15,16
		/// </summary>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public event FormOldV10_BeforeDelConfirmEventHandler BeforeDelConfirmEvent
		{
			add
			{
				CreateEventBridge();
				_BeforeDelConfirmEvent += value;
			}
			remove
			{
				_BeforeDelConfirmEvent -= value;
			}
		}

		/// <summary>
		/// SupportByVersion Access, 9,10,11,12,14,15,16
		/// </summary>
		private event FormOldV10_AfterDelConfirmEventHandler _AfterDelConfirmEvent;

		/// <summary>
		/// SupportByVersion Access 9 10 11 12 14 15,16
		/// </summary>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public event FormOldV10_AfterDelConfirmEventHandler AfterDelConfirmEvent
		{
			add
			{
				CreateEventBridge();
				_AfterDelConfirmEvent += value;
			}
			remove
			{
				_AfterDelConfirmEvent -= value;
			}
		}

		/// <summary>
		/// SupportByVersion Access, 9,10,11,12,14,15,16
		/// </summary>
		private event FormOldV10_OpenEventHandler _OpenEvent;

		/// <summary>
		/// SupportByVersion Access 9 10 11 12 14 15,16
		/// </summary>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public event FormOldV10_OpenEventHandler OpenEvent
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
		/// SupportByVersion Access, 9,10,11,12,14,15,16
		/// </summary>
		private event FormOldV10_ResizeEventHandler _ResizeEvent;

		/// <summary>
		/// SupportByVersion Access 9 10 11 12 14 15,16
		/// </summary>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public event FormOldV10_ResizeEventHandler ResizeEvent
		{
			add
			{
				CreateEventBridge();
				_ResizeEvent += value;
			}
			remove
			{
				_ResizeEvent -= value;
			}
		}

		/// <summary>
		/// SupportByVersion Access, 9,10,11,12,14,15,16
		/// </summary>
		private event FormOldV10_UnloadEventHandler _UnloadEvent;

		/// <summary>
		/// SupportByVersion Access 9 10 11 12 14 15,16
		/// </summary>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public event FormOldV10_UnloadEventHandler UnloadEvent
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
		/// SupportByVersion Access, 9,10,11,12,14,15,16
		/// </summary>
		private event FormOldV10_CloseEventHandler _CloseEvent;

		/// <summary>
		/// SupportByVersion Access 9 10 11 12 14 15,16
		/// </summary>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public event FormOldV10_CloseEventHandler CloseEvent
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
		/// SupportByVersion Access, 9,10,11,12,14,15,16
		/// </summary>
		private event FormOldV10_ActivateEventHandler _ActivateEvent;

		/// <summary>
		/// SupportByVersion Access 9 10 11 12 14 15,16
		/// </summary>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public event FormOldV10_ActivateEventHandler ActivateEvent
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
		/// SupportByVersion Access, 9,10,11,12,14,15,16
		/// </summary>
		private event FormOldV10_DeactivateEventHandler _DeactivateEvent;

		/// <summary>
		/// SupportByVersion Access 9 10 11 12 14 15,16
		/// </summary>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public event FormOldV10_DeactivateEventHandler DeactivateEvent
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
		/// SupportByVersion Access, 9,10,11,12,14,15,16
		/// </summary>
		private event FormOldV10_GotFocusEventHandler _GotFocusEvent;

		/// <summary>
		/// SupportByVersion Access 9 10 11 12 14 15,16
		/// </summary>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public event FormOldV10_GotFocusEventHandler GotFocusEvent
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
		/// SupportByVersion Access, 9,10,11,12,14,15,16
		/// </summary>
		private event FormOldV10_LostFocusEventHandler _LostFocusEvent;

		/// <summary>
		/// SupportByVersion Access 9 10 11 12 14 15,16
		/// </summary>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public event FormOldV10_LostFocusEventHandler LostFocusEvent
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
		/// SupportByVersion Access, 9,10,11,12,14,15,16
		/// </summary>
		private event FormOldV10_ClickEventHandler _ClickEvent;

		/// <summary>
		/// SupportByVersion Access 9 10 11 12 14 15,16
		/// </summary>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public event FormOldV10_ClickEventHandler ClickEvent
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
		/// SupportByVersion Access, 9,10,11,12,14,15,16
		/// </summary>
		private event FormOldV10_DblClickEventHandler _DblClickEvent;

		/// <summary>
		/// SupportByVersion Access 9 10 11 12 14 15,16
		/// </summary>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public event FormOldV10_DblClickEventHandler DblClickEvent
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
		/// SupportByVersion Access, 9,10,11,12,14,15,16
		/// </summary>
		private event FormOldV10_MouseDownEventHandler _MouseDownEvent;

		/// <summary>
		/// SupportByVersion Access 9 10 11 12 14 15,16
		/// </summary>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public event FormOldV10_MouseDownEventHandler MouseDownEvent
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
		/// SupportByVersion Access, 9,10,11,12,14,15,16
		/// </summary>
		private event FormOldV10_MouseMoveEventHandler _MouseMoveEvent;

		/// <summary>
		/// SupportByVersion Access 9 10 11 12 14 15,16
		/// </summary>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public event FormOldV10_MouseMoveEventHandler MouseMoveEvent
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
		/// SupportByVersion Access, 9,10,11,12,14,15,16
		/// </summary>
		private event FormOldV10_MouseUpEventHandler _MouseUpEvent;

		/// <summary>
		/// SupportByVersion Access 9 10 11 12 14 15,16
		/// </summary>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public event FormOldV10_MouseUpEventHandler MouseUpEvent
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
		/// SupportByVersion Access, 9,10,11,12,14,15,16
		/// </summary>
		private event FormOldV10_KeyDownEventHandler _KeyDownEvent;

		/// <summary>
		/// SupportByVersion Access 9 10 11 12 14 15,16
		/// </summary>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public event FormOldV10_KeyDownEventHandler KeyDownEvent
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
		/// SupportByVersion Access, 9,10,11,12,14,15,16
		/// </summary>
		private event FormOldV10_KeyPressEventHandler _KeyPressEvent;

		/// <summary>
		/// SupportByVersion Access 9 10 11 12 14 15,16
		/// </summary>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public event FormOldV10_KeyPressEventHandler KeyPressEvent
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
		/// SupportByVersion Access, 9,10,11,12,14,15,16
		/// </summary>
		private event FormOldV10_KeyUpEventHandler _KeyUpEvent;

		/// <summary>
		/// SupportByVersion Access 9 10 11 12 14 15,16
		/// </summary>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public event FormOldV10_KeyUpEventHandler KeyUpEvent
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
		/// SupportByVersion Access, 9,10,11,12,14,15,16
		/// </summary>
		private event FormOldV10_ErrorEventHandler _ErrorEvent;

		/// <summary>
		/// SupportByVersion Access 9 10 11 12 14 15,16
		/// </summary>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public event FormOldV10_ErrorEventHandler ErrorEvent
		{
			add
			{
				CreateEventBridge();
				_ErrorEvent += value;
			}
			remove
			{
				_ErrorEvent -= value;
			}
		}

		/// <summary>
		/// SupportByVersion Access, 9,10,11,12,14,15,16
		/// </summary>
		private event FormOldV10_TimerEventHandler _TimerEvent;

		/// <summary>
		/// SupportByVersion Access 9 10 11 12 14 15,16
		/// </summary>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public event FormOldV10_TimerEventHandler TimerEvent
		{
			add
			{
				CreateEventBridge();
				_TimerEvent += value;
			}
			remove
			{
				_TimerEvent -= value;
			}
		}

		/// <summary>
		/// SupportByVersion Access, 9,10,11,12,14,15,16
		/// </summary>
		private event FormOldV10_FilterEventHandler _FilterEvent;

		/// <summary>
		/// SupportByVersion Access 9 10 11 12 14 15,16
		/// </summary>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public event FormOldV10_FilterEventHandler FilterEvent
		{
			add
			{
				CreateEventBridge();
				_FilterEvent += value;
			}
			remove
			{
				_FilterEvent -= value;
			}
		}

		/// <summary>
		/// SupportByVersion Access, 9,10,11,12,14,15,16
		/// </summary>
		private event FormOldV10_ApplyFilterEventHandler _ApplyFilterEvent;

		/// <summary>
		/// SupportByVersion Access 9 10 11 12 14 15,16
		/// </summary>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public event FormOldV10_ApplyFilterEventHandler ApplyFilterEvent
		{
			add
			{
				CreateEventBridge();
				_ApplyFilterEvent += value;
			}
			remove
			{
				_ApplyFilterEvent -= value;
			}
		}

		/// <summary>
		/// SupportByVersion Access, 9,10,11,12,14,15,16
		/// </summary>
		private event FormOldV10_DirtyEventHandler _DirtyEvent;

		/// <summary>
		/// SupportByVersion Access 9 10 11 12 14 15,16
		/// </summary>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public event FormOldV10_DirtyEventHandler DirtyEvent
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
		/// SupportByVersion Access, 10,11,12,14,15,16
		/// </summary>
		private event FormOldV10_UndoEventHandler _UndoEvent;

		/// <summary>
		/// SupportByVersion Access 10 11 12 14 15,16
		/// </summary>
		[SupportByVersion("Access", 10,11,12,14,15,16)]
		public event FormOldV10_UndoEventHandler UndoEvent
		{
			add
			{
				CreateEventBridge();
				_UndoEvent += value;
			}
			remove
			{
				_UndoEvent -= value;
			}
		}

		/// <summary>
		/// SupportByVersion Access, 10,11,12,14,15,16
		/// </summary>
		private event FormOldV10_RecordExitEventHandler _RecordExitEvent;

		/// <summary>
		/// SupportByVersion Access 10 11 12 14 15,16
		/// </summary>
		[SupportByVersion("Access", 10,11,12,14,15,16)]
		public event FormOldV10_RecordExitEventHandler RecordExitEvent
		{
			add
			{
				CreateEventBridge();
				_RecordExitEvent += value;
			}
			remove
			{
				_RecordExitEvent -= value;
			}
		}

		/// <summary>
		/// SupportByVersion Access, 10,11,12,14,15,16
		/// </summary>
		private event FormOldV10_BeginBatchEditEventHandler _BeginBatchEditEvent;

		/// <summary>
		/// SupportByVersion Access 10 11 12 14 15,16
		/// </summary>
		[SupportByVersion("Access", 10,11,12,14,15,16)]
		public event FormOldV10_BeginBatchEditEventHandler BeginBatchEditEvent
		{
			add
			{
				CreateEventBridge();
				_BeginBatchEditEvent += value;
			}
			remove
			{
				_BeginBatchEditEvent -= value;
			}
		}

		/// <summary>
		/// SupportByVersion Access, 10,11,12,14,15,16
		/// </summary>
		private event FormOldV10_UndoBatchEditEventHandler _UndoBatchEditEvent;

		/// <summary>
		/// SupportByVersion Access 10 11 12 14 15,16
		/// </summary>
		[SupportByVersion("Access", 10,11,12,14,15,16)]
		public event FormOldV10_UndoBatchEditEventHandler UndoBatchEditEvent
		{
			add
			{
				CreateEventBridge();
				_UndoBatchEditEvent += value;
			}
			remove
			{
				_UndoBatchEditEvent -= value;
			}
		}

		/// <summary>
		/// SupportByVersion Access, 10,11,12,14,15,16
		/// </summary>
		private event FormOldV10_BeforeBeginTransactionEventHandler _BeforeBeginTransactionEvent;

		/// <summary>
		/// SupportByVersion Access 10 11 12 14 15,16
		/// </summary>
		[SupportByVersion("Access", 10,11,12,14,15,16)]
		public event FormOldV10_BeforeBeginTransactionEventHandler BeforeBeginTransactionEvent
		{
			add
			{
				CreateEventBridge();
				_BeforeBeginTransactionEvent += value;
			}
			remove
			{
				_BeforeBeginTransactionEvent -= value;
			}
		}

		/// <summary>
		/// SupportByVersion Access, 10,11,12,14,15,16
		/// </summary>
		private event FormOldV10_AfterBeginTransactionEventHandler _AfterBeginTransactionEvent;

		/// <summary>
		/// SupportByVersion Access 10 11 12 14 15,16
		/// </summary>
		[SupportByVersion("Access", 10,11,12,14,15,16)]
		public event FormOldV10_AfterBeginTransactionEventHandler AfterBeginTransactionEvent
		{
			add
			{
				CreateEventBridge();
				_AfterBeginTransactionEvent += value;
			}
			remove
			{
				_AfterBeginTransactionEvent -= value;
			}
		}

		/// <summary>
		/// SupportByVersion Access, 10,11,12,14,15,16
		/// </summary>
		private event FormOldV10_BeforeCommitTransactionEventHandler _BeforeCommitTransactionEvent;

		/// <summary>
		/// SupportByVersion Access 10 11 12 14 15,16
		/// </summary>
		[SupportByVersion("Access", 10,11,12,14,15,16)]
		public event FormOldV10_BeforeCommitTransactionEventHandler BeforeCommitTransactionEvent
		{
			add
			{
				CreateEventBridge();
				_BeforeCommitTransactionEvent += value;
			}
			remove
			{
				_BeforeCommitTransactionEvent -= value;
			}
		}

		/// <summary>
		/// SupportByVersion Access, 10,11,12,14,15,16
		/// </summary>
		private event FormOldV10_AfterCommitTransactionEventHandler _AfterCommitTransactionEvent;

		/// <summary>
		/// SupportByVersion Access 10 11 12 14 15,16
		/// </summary>
		[SupportByVersion("Access", 10,11,12,14,15,16)]
		public event FormOldV10_AfterCommitTransactionEventHandler AfterCommitTransactionEvent
		{
			add
			{
				CreateEventBridge();
				_AfterCommitTransactionEvent += value;
			}
			remove
			{
				_AfterCommitTransactionEvent -= value;
			}
		}

		/// <summary>
		/// SupportByVersion Access, 10,11,12,14,15,16
		/// </summary>
		private event FormOldV10_RollbackTransactionEventHandler _RollbackTransactionEvent;

		/// <summary>
		/// SupportByVersion Access 10 11 12 14 15,16
		/// </summary>
		[SupportByVersion("Access", 10,11,12,14,15,16)]
		public event FormOldV10_RollbackTransactionEventHandler RollbackTransactionEvent
		{
			add
			{
				CreateEventBridge();
				_RollbackTransactionEvent += value;
			}
			remove
			{
				_RollbackTransactionEvent -= value;
			}
		}

		/// <summary>
		/// SupportByVersion Access, 10,11,12,14,15,16
		/// </summary>
		private event FormOldV10_OnConnectEventHandler _OnConnectEvent;

		/// <summary>
		/// SupportByVersion Access 10 11 12 14 15,16
		/// </summary>
		[SupportByVersion("Access", 10,11,12,14,15,16)]
		public event FormOldV10_OnConnectEventHandler OnConnectEvent
		{
			add
			{
				CreateEventBridge();
				_OnConnectEvent += value;
			}
			remove
			{
				_OnConnectEvent -= value;
			}
		}

		/// <summary>
		/// SupportByVersion Access, 10,11,12,14,15,16
		/// </summary>
		private event FormOldV10_OnDisconnectEventHandler _OnDisconnectEvent;

		/// <summary>
		/// SupportByVersion Access 10 11 12 14 15,16
		/// </summary>
		[SupportByVersion("Access", 10,11,12,14,15,16)]
		public event FormOldV10_OnDisconnectEventHandler OnDisconnectEvent
		{
			add
			{
				CreateEventBridge();
				_OnDisconnectEvent += value;
			}
			remove
			{
				_OnDisconnectEvent -= value;
			}
		}

		/// <summary>
		/// SupportByVersion Access, 10,11,12,14,15,16
		/// </summary>
		private event FormOldV10_PivotTableChangeEventHandler _PivotTableChangeEvent;

		/// <summary>
		/// SupportByVersion Access 10 11 12 14 15,16
		/// </summary>
		[SupportByVersion("Access", 10,11,12,14,15,16)]
		public event FormOldV10_PivotTableChangeEventHandler PivotTableChangeEvent
		{
			add
			{
				CreateEventBridge();
				_PivotTableChangeEvent += value;
			}
			remove
			{
				_PivotTableChangeEvent -= value;
			}
		}

		/// <summary>
		/// SupportByVersion Access, 10,11,12,14,15,16
		/// </summary>
		private event FormOldV10_QueryEventHandler _QueryEvent;

		/// <summary>
		/// SupportByVersion Access 10 11 12 14 15,16
		/// </summary>
		[SupportByVersion("Access", 10,11,12,14,15,16)]
		public event FormOldV10_QueryEventHandler QueryEvent
		{
			add
			{
				CreateEventBridge();
				_QueryEvent += value;
			}
			remove
			{
				_QueryEvent -= value;
			}
		}

		/// <summary>
		/// SupportByVersion Access, 10,11,12,14,15,16
		/// </summary>
		private event FormOldV10_BeforeQueryEventHandler _BeforeQueryEvent;

		/// <summary>
		/// SupportByVersion Access 10 11 12 14 15,16
		/// </summary>
		[SupportByVersion("Access", 10,11,12,14,15,16)]
		public event FormOldV10_BeforeQueryEventHandler BeforeQueryEvent
		{
			add
			{
				CreateEventBridge();
				_BeforeQueryEvent += value;
			}
			remove
			{
				_BeforeQueryEvent -= value;
			}
		}

		/// <summary>
		/// SupportByVersion Access, 10,11,12,14,15,16
		/// </summary>
		private event FormOldV10_SelectionChangeEventHandler _SelectionChangeEvent;

		/// <summary>
		/// SupportByVersion Access 10 11 12 14 15,16
		/// </summary>
		[SupportByVersion("Access", 10,11,12,14,15,16)]
		public event FormOldV10_SelectionChangeEventHandler SelectionChangeEvent
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
		/// SupportByVersion Access, 10,11,12,14,15,16
		/// </summary>
		private event FormOldV10_CommandBeforeExecuteEventHandler _CommandBeforeExecuteEvent;

		/// <summary>
		/// SupportByVersion Access 10 11 12 14 15,16
		/// </summary>
		[SupportByVersion("Access", 10,11,12,14,15,16)]
		public event FormOldV10_CommandBeforeExecuteEventHandler CommandBeforeExecuteEvent
		{
			add
			{
				CreateEventBridge();
				_CommandBeforeExecuteEvent += value;
			}
			remove
			{
				_CommandBeforeExecuteEvent -= value;
			}
		}

		/// <summary>
		/// SupportByVersion Access, 10,11,12,14,15,16
		/// </summary>
		private event FormOldV10_CommandCheckedEventHandler _CommandCheckedEvent;

		/// <summary>
		/// SupportByVersion Access 10 11 12 14 15,16
		/// </summary>
		[SupportByVersion("Access", 10,11,12,14,15,16)]
		public event FormOldV10_CommandCheckedEventHandler CommandCheckedEvent
		{
			add
			{
				CreateEventBridge();
				_CommandCheckedEvent += value;
			}
			remove
			{
				_CommandCheckedEvent -= value;
			}
		}

		/// <summary>
		/// SupportByVersion Access, 10,11,12,14,15,16
		/// </summary>
		private event FormOldV10_CommandEnabledEventHandler _CommandEnabledEvent;

		/// <summary>
		/// SupportByVersion Access 10 11 12 14 15,16
		/// </summary>
		[SupportByVersion("Access", 10,11,12,14,15,16)]
		public event FormOldV10_CommandEnabledEventHandler CommandEnabledEvent
		{
			add
			{
				CreateEventBridge();
				_CommandEnabledEvent += value;
			}
			remove
			{
				_CommandEnabledEvent -= value;
			}
		}

		/// <summary>
		/// SupportByVersion Access, 10,11,12,14,15,16
		/// </summary>
		private event FormOldV10_CommandExecuteEventHandler _CommandExecuteEvent;

		/// <summary>
		/// SupportByVersion Access 10 11 12 14 15,16
		/// </summary>
		[SupportByVersion("Access", 10,11,12,14,15,16)]
		public event FormOldV10_CommandExecuteEventHandler CommandExecuteEvent
		{
			add
			{
				CreateEventBridge();
				_CommandExecuteEvent += value;
			}
			remove
			{
				_CommandExecuteEvent -= value;
			}
		}

		/// <summary>
		/// SupportByVersion Access, 10,11,12,14,15,16
		/// </summary>
		private event FormOldV10_DataSetChangeEventHandler _DataSetChangeEvent;

		/// <summary>
		/// SupportByVersion Access 10 11 12 14 15,16
		/// </summary>
		[SupportByVersion("Access", 10,11,12,14,15,16)]
		public event FormOldV10_DataSetChangeEventHandler DataSetChangeEvent
		{
			add
			{
				CreateEventBridge();
				_DataSetChangeEvent += value;
			}
			remove
			{
				_DataSetChangeEvent -= value;
			}
		}

		/// <summary>
		/// SupportByVersion Access, 10,11,12,14,15,16
		/// </summary>
		private event FormOldV10_BeforeScreenTipEventHandler _BeforeScreenTipEvent;

		/// <summary>
		/// SupportByVersion Access 10 11 12 14 15,16
		/// </summary>
		[SupportByVersion("Access", 10,11,12,14,15,16)]
		public event FormOldV10_BeforeScreenTipEventHandler BeforeScreenTipEvent
		{
			add
			{
				CreateEventBridge();
				_BeforeScreenTipEvent += value;
			}
			remove
			{
				_BeforeScreenTipEvent -= value;
			}
		}

		/// <summary>
		/// SupportByVersion Access, 10,11,12,14,15,16
		/// </summary>
		private event FormOldV10_BeforeRenderEventHandler _BeforeRenderEvent;

		/// <summary>
		/// SupportByVersion Access 10 11 12 14 15,16
		/// </summary>
		[SupportByVersion("Access", 10,11,12,14,15,16)]
		public event FormOldV10_BeforeRenderEventHandler BeforeRenderEvent
		{
			add
			{
				CreateEventBridge();
				_BeforeRenderEvent += value;
			}
			remove
			{
				_BeforeRenderEvent -= value;
			}
		}

		/// <summary>
		/// SupportByVersion Access, 10,11,12,14,15,16
		/// </summary>
		private event FormOldV10_AfterRenderEventHandler _AfterRenderEvent;

		/// <summary>
		/// SupportByVersion Access 10 11 12 14 15,16
		/// </summary>
		[SupportByVersion("Access", 10,11,12,14,15,16)]
		public event FormOldV10_AfterRenderEventHandler AfterRenderEvent
		{
			add
			{
				CreateEventBridge();
				_AfterRenderEvent += value;
			}
			remove
			{
				_AfterRenderEvent -= value;
			}
		}

		/// <summary>
		/// SupportByVersion Access, 10,11,12,14,15,16
		/// </summary>
		private event FormOldV10_AfterFinalRenderEventHandler _AfterFinalRenderEvent;

		/// <summary>
		/// SupportByVersion Access 10 11 12 14 15,16
		/// </summary>
		[SupportByVersion("Access", 10,11,12,14,15,16)]
		public event FormOldV10_AfterFinalRenderEventHandler AfterFinalRenderEvent
		{
			add
			{
				CreateEventBridge();
				_AfterFinalRenderEvent += value;
			}
			remove
			{
				_AfterFinalRenderEvent -= value;
			}
		}

		/// <summary>
		/// SupportByVersion Access, 10,11,12,14,15,16
		/// </summary>
		private event FormOldV10_AfterLayoutEventHandler _AfterLayoutEvent;

		/// <summary>
		/// SupportByVersion Access 10 11 12 14 15,16
		/// </summary>
		[SupportByVersion("Access", 10,11,12,14,15,16)]
		public event FormOldV10_AfterLayoutEventHandler AfterLayoutEvent
		{
			add
			{
				CreateEventBridge();
				_AfterLayoutEvent += value;
			}
			remove
			{
				_AfterLayoutEvent -= value;
			}
		}

		/// <summary>
		/// SupportByVersion Access, 10,11,12,14,15,16
		/// </summary>
		private event FormOldV10_MouseWheelEventHandler _MouseWheelEvent;

		/// <summary>
		/// SupportByVersion Access 10 11 12 14 15,16
		/// </summary>
		[SupportByVersion("Access", 10,11,12,14,15,16)]
		public event FormOldV10_MouseWheelEventHandler MouseWheelEvent
		{
			add
			{
				CreateEventBridge();
				_MouseWheelEvent += value;
			}
			remove
			{
				_MouseWheelEvent -= value;
			}
		}

		/// <summary>
		/// SupportByVersion Access, 10,11,12,14,15,16
		/// </summary>
		private event FormOldV10_ViewChangeEventHandler _ViewChangeEvent;

		/// <summary>
		/// SupportByVersion Access 10 11 12 14 15,16
		/// </summary>
		[SupportByVersion("Access", 10,11,12,14,15,16)]
		public event FormOldV10_ViewChangeEventHandler ViewChangeEvent
		{
			add
			{
				CreateEventBridge();
				_ViewChangeEvent += value;
			}
			remove
			{
				_ViewChangeEvent -= value;
			}
		}

		/// <summary>
		/// SupportByVersion Access, 10,11,12,14,15,16
		/// </summary>
		private event FormOldV10_DataChangeEventHandler _DataChangeEvent;

		/// <summary>
		/// SupportByVersion Access 10 11 12 14 15,16
		/// </summary>
		[SupportByVersion("Access", 10,11,12,14,15,16)]
		public event FormOldV10_DataChangeEventHandler DataChangeEvent
		{
			add
			{
				CreateEventBridge();
				_DataChangeEvent += value;
			}
			remove
			{
				_DataChangeEvent -= value;
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
				_activeSinkId = SinkHelper.GetConnectionPoint(this, ref _connectPoint, Events._FormEvents_SinkHelper.Id);


			if(Events._FormEvents_SinkHelper.Id.Equals(_activeSinkId, StringComparison.InvariantCultureIgnoreCase))
			{
				__FormEvents_SinkHelper = new Events._FormEvents_SinkHelper(this, _connectPoint);
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
			if( null != __FormEvents_SinkHelper)
			{
				__FormEvents_SinkHelper.Dispose();
				__FormEvents_SinkHelper = null;
			}

			_connectPoint = null;
		}
        
        #endregion

		#pragma warning restore
	}
}

