using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.AccessApi
{
	#region Delegates

	#pragma warning disable
	public delegate void FormOld_LoadEventHandler();
	public delegate void FormOld_CurrentEventHandler();
	public delegate void FormOld_BeforeInsertEventHandler(ref Int16 cancel);
	public delegate void FormOld_AfterInsertEventHandler();
	public delegate void FormOld_BeforeUpdateEventHandler(ref Int16 cancel);
	public delegate void FormOld_AfterUpdateEventHandler();
	public delegate void FormOld_DeleteEventHandler(ref Int16 cancel);
	public delegate void FormOld_BeforeDelConfirmEventHandler(ref Int16 cancel, ref Int16 response);
	public delegate void FormOld_AfterDelConfirmEventHandler(ref Int16 status);
	public delegate void FormOld_OpenEventHandler(ref Int16 cancel);
	public delegate void FormOld_ResizeEventHandler();
	public delegate void FormOld_UnloadEventHandler(ref Int16 cancel);
	public delegate void FormOld_CloseEventHandler();
	public delegate void FormOld_ActivateEventHandler();
	public delegate void FormOld_DeactivateEventHandler();
	public delegate void FormOld_GotFocusEventHandler();
	public delegate void FormOld_LostFocusEventHandler();
	public delegate void FormOld_ClickEventHandler();
	public delegate void FormOld_DblClickEventHandler(ref Int16 cancel);
	public delegate void FormOld_MouseDownEventHandler(ref Int16 button, ref Int16 shift, ref Single x, ref Single y);
	public delegate void FormOld_MouseMoveEventHandler(ref Int16 button, ref Int16 shift, ref Single x, ref Single y);
	public delegate void FormOld_MouseUpEventHandler(ref Int16 button, ref Int16 shift, ref Single x, ref Single y);
	public delegate void FormOld_KeyDownEventHandler(ref Int16 keyCode, ref Int16 shift);
	public delegate void FormOld_KeyPressEventHandler(ref Int16 keyAscii);
	public delegate void FormOld_KeyUpEventHandler(ref Int16 keyCode, ref Int16 shift);
	public delegate void FormOld_ErrorEventHandler(ref Int16 dataErr, ref Int16 response);
	public delegate void FormOld_TimerEventHandler();
	public delegate void FormOld_FilterEventHandler(ref Int16 cancel, ref Int16 filterType);
	public delegate void FormOld_ApplyFilterEventHandler(ref Int16 cancel, ref Int16 applyType);
	public delegate void FormOld_DirtyEventHandler(ref Int16 cancel);
	public delegate void FormOld_UndoEventHandler(ref Int16 cancel);
	public delegate void FormOld_RecordExitEventHandler(ref Int16 cancel);
	public delegate void FormOld_BeginBatchEditEventHandler(ref Int16 cancel);
	public delegate void FormOld_UndoBatchEditEventHandler(ref Int16 cancel);
	public delegate void FormOld_BeforeBeginTransactionEventHandler(ref Int16 cancel, ref NetOffice.ADODBApi.Connection connection);
	public delegate void FormOld_AfterBeginTransactionEventHandler(ref NetOffice.ADODBApi.Connection connection);
	public delegate void FormOld_BeforeCommitTransactionEventHandler(ref Int16 cancel, ref NetOffice.ADODBApi.Connection connection);
	public delegate void FormOld_AfterCommitTransactionEventHandler(ref NetOffice.ADODBApi.Connection connection);
	public delegate void FormOld_RollbackTransactionEventHandler(ref NetOffice.ADODBApi.Connection connection);
	public delegate void FormOld_OnConnectEventHandler();
	public delegate void FormOld_OnDisconnectEventHandler();
	public delegate void FormOld_PivotTableChangeEventHandler(Int32 reason);
	public delegate void FormOld_QueryEventHandler();
	public delegate void FormOld_BeforeQueryEventHandler();
	public delegate void FormOld_SelectionChangeEventHandler();
	public delegate void FormOld_CommandBeforeExecuteEventHandler(object command, ICOMObject cancel);
	public delegate void FormOld_CommandCheckedEventHandler(object command, ICOMObject Checked);
	public delegate void FormOld_CommandEnabledEventHandler(object command, ICOMObject enabled);
	public delegate void FormOld_CommandExecuteEventHandler(object command);
	public delegate void FormOld_DataSetChangeEventHandler();
	public delegate void FormOld_BeforeScreenTipEventHandler(ICOMObject screenTipText, ICOMObject sourceObject);
	public delegate void FormOld_BeforeRenderEventHandler(ICOMObject drawObject, ICOMObject chartObject, ICOMObject cancel);
	public delegate void FormOld_AfterRenderEventHandler(ICOMObject drawObject, ICOMObject chartObject);
	public delegate void FormOld_AfterFinalRenderEventHandler(ICOMObject drawObject);
	public delegate void FormOld_AfterLayoutEventHandler(ICOMObject drawObject);
	public delegate void FormOld_MouseWheelEventHandler(bool page, Int32 count);
	public delegate void FormOld_ViewChangeEventHandler(Int32 reason);
	public delegate void FormOld_DataChangeEventHandler(Int32 reason);
	#pragma warning restore

	#endregion

	/// <summary>
	/// CoClass FormOld 
	/// SupportByVersion Access, 10,11,12,14,15,16
	/// </summary>
	[SupportByVersion("Access", 10,11,12,14,15,16)]
	[EntityType(EntityType.IsCoClass)]
    [ComEventContract(typeof(EventContracts._FormEvents))]
	[TypeId("483615A0-74BE-101B-AF4E-00AA003F0F07")]
    public interface FormOld : _Form, IEventBinding
	{
		#region Events

		/// <summary>
		/// SupportByVersion Access 9 10 11 12 14 15,16
		/// </summary>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		event FormOld_LoadEventHandler LoadEvent;

		/// <summary>
		/// SupportByVersion Access 9 10 11 12 14 15,16
		/// </summary>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		event FormOld_CurrentEventHandler CurrentEvent;

		/// <summary>
		/// SupportByVersion Access 9 10 11 12 14 15,16
		/// </summary>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		event FormOld_BeforeInsertEventHandler BeforeInsertEvent;

		/// <summary>
		/// SupportByVersion Access 9 10 11 12 14 15,16
		/// </summary>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		event FormOld_AfterInsertEventHandler AfterInsertEvent;

		/// <summary>
		/// SupportByVersion Access 9 10 11 12 14 15,16
		/// </summary>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		event FormOld_BeforeUpdateEventHandler BeforeUpdateEvent;

		/// <summary>
		/// SupportByVersion Access 9 10 11 12 14 15,16
		/// </summary>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		event FormOld_AfterUpdateEventHandler AfterUpdateEvent;

		/// <summary>
		/// SupportByVersion Access 9 10 11 12 14 15,16
		/// </summary>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		event FormOld_DeleteEventHandler DeleteEvent;

		/// <summary>
		/// SupportByVersion Access 9 10 11 12 14 15,16
		/// </summary>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		event FormOld_BeforeDelConfirmEventHandler BeforeDelConfirmEvent;

		/// <summary>
		/// SupportByVersion Access 9 10 11 12 14 15,16
		/// </summary>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		event FormOld_AfterDelConfirmEventHandler AfterDelConfirmEvent;

		/// <summary>
		/// SupportByVersion Access 9 10 11 12 14 15,16
		/// </summary>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		event FormOld_OpenEventHandler OpenEvent;

		/// <summary>
		/// SupportByVersion Access 9 10 11 12 14 15,16
		/// </summary>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		event FormOld_ResizeEventHandler ResizeEvent;

		/// <summary>
		/// SupportByVersion Access 9 10 11 12 14 15,16
		/// </summary>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		event FormOld_UnloadEventHandler UnloadEvent;

		/// <summary>
		/// SupportByVersion Access 9 10 11 12 14 15,16
		/// </summary>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		event FormOld_CloseEventHandler CloseEvent;

		/// <summary>
		/// SupportByVersion Access 9 10 11 12 14 15,16
		/// </summary>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		event FormOld_ActivateEventHandler ActivateEvent;

		/// <summary>
		/// SupportByVersion Access 9 10 11 12 14 15,16
		/// </summary>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		event FormOld_DeactivateEventHandler DeactivateEvent;

		/// <summary>
		/// SupportByVersion Access 9 10 11 12 14 15,16
		/// </summary>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		event FormOld_GotFocusEventHandler GotFocusEvent;

		/// <summary>
		/// SupportByVersion Access 9 10 11 12 14 15,16
		/// </summary>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		event FormOld_LostFocusEventHandler LostFocusEvent;

		/// <summary>
		/// SupportByVersion Access 9 10 11 12 14 15,16
		/// </summary>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		event FormOld_ClickEventHandler ClickEvent;

		/// <summary>
		/// SupportByVersion Access 9 10 11 12 14 15,16
		/// </summary>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		event FormOld_DblClickEventHandler DblClickEvent;

		/// <summary>
		/// SupportByVersion Access 9 10 11 12 14 15,16
		/// </summary>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		event FormOld_MouseDownEventHandler MouseDownEvent;

		/// <summary>
		/// SupportByVersion Access 9 10 11 12 14 15,16
		/// </summary>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		event FormOld_MouseMoveEventHandler MouseMoveEvent;

		/// <summary>
		/// SupportByVersion Access 9 10 11 12 14 15,16
		/// </summary>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		event FormOld_MouseUpEventHandler MouseUpEvent;

		/// <summary>
		/// SupportByVersion Access 9 10 11 12 14 15,16
		/// </summary>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		event FormOld_KeyDownEventHandler KeyDownEvent;

		/// <summary>
		/// SupportByVersion Access 9 10 11 12 14 15,16
		/// </summary>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		event FormOld_KeyPressEventHandler KeyPressEvent;

		/// <summary>
		/// SupportByVersion Access 9 10 11 12 14 15,16
		/// </summary>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		event FormOld_KeyUpEventHandler KeyUpEvent;

		/// <summary>
		/// SupportByVersion Access 9 10 11 12 14 15,16
		/// </summary>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		event FormOld_ErrorEventHandler ErrorEvent;

		/// <summary>
		/// SupportByVersion Access 9 10 11 12 14 15,16
		/// </summary>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		event FormOld_TimerEventHandler TimerEvent;

		/// <summary>
		/// SupportByVersion Access 9 10 11 12 14 15,16
		/// </summary>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		event FormOld_FilterEventHandler FilterEvent;

		/// <summary>
		/// SupportByVersion Access 9 10 11 12 14 15,16
		/// </summary>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		event FormOld_ApplyFilterEventHandler ApplyFilterEvent;

		/// <summary>
		/// SupportByVersion Access 9 10 11 12 14 15,16
		/// </summary>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		event FormOld_DirtyEventHandler DirtyEvent;

		/// <summary>
		/// SupportByVersion Access 10 11 12 14 15,16
		/// </summary>
		[SupportByVersion("Access", 10,11,12,14,15,16)]
		event FormOld_UndoEventHandler UndoEvent;

		/// <summary>
		/// SupportByVersion Access 10 11 12 14 15,16
		/// </summary>
		[SupportByVersion("Access", 10,11,12,14,15,16)]
		event FormOld_RecordExitEventHandler RecordExitEvent;

		/// <summary>
		/// SupportByVersion Access 10 11 12 14 15,16
		/// </summary>
		[SupportByVersion("Access", 10,11,12,14,15,16)]
		event FormOld_BeginBatchEditEventHandler BeginBatchEditEvent;

		/// <summary>
		/// SupportByVersion Access 10 11 12 14 15,16
		/// </summary>
		[SupportByVersion("Access", 10,11,12,14,15,16)]
		event FormOld_UndoBatchEditEventHandler UndoBatchEditEvent;

		/// <summary>
		/// SupportByVersion Access 10 11 12 14 15,16
		/// </summary>
		[SupportByVersion("Access", 10,11,12,14,15,16)]
		event FormOld_BeforeBeginTransactionEventHandler BeforeBeginTransactionEvent;

		/// <summary>
		/// SupportByVersion Access 10 11 12 14 15,16
		/// </summary>
		[SupportByVersion("Access", 10,11,12,14,15,16)]
		event FormOld_AfterBeginTransactionEventHandler AfterBeginTransactionEvent;

		/// <summary>
		/// SupportByVersion Access 10 11 12 14 15,16
		/// </summary>
		[SupportByVersion("Access", 10,11,12,14,15,16)]
		event FormOld_BeforeCommitTransactionEventHandler BeforeCommitTransactionEvent;

		/// <summary>
		/// SupportByVersion Access 10 11 12 14 15,16
		/// </summary>
		[SupportByVersion("Access", 10,11,12,14,15,16)]
		event FormOld_AfterCommitTransactionEventHandler AfterCommitTransactionEvent;

		/// <summary>
		/// SupportByVersion Access 10 11 12 14 15,16
		/// </summary>
		[SupportByVersion("Access", 10,11,12,14,15,16)]
		event FormOld_RollbackTransactionEventHandler RollbackTransactionEvent;

		/// <summary>
		/// SupportByVersion Access 10 11 12 14 15,16
		/// </summary>
		[SupportByVersion("Access", 10,11,12,14,15,16)]
		event FormOld_OnConnectEventHandler OnConnectEvent;

		/// <summary>
		/// SupportByVersion Access 10 11 12 14 15,16
		/// </summary>
		[SupportByVersion("Access", 10,11,12,14,15,16)]
		event FormOld_OnDisconnectEventHandler OnDisconnectEvent;

		/// <summary>
		/// SupportByVersion Access 10 11 12 14 15,16
		/// </summary>
		[SupportByVersion("Access", 10,11,12,14,15,16)]
		event FormOld_PivotTableChangeEventHandler PivotTableChangeEvent;

		/// <summary>
		/// SupportByVersion Access 10 11 12 14 15,16
		/// </summary>
		[SupportByVersion("Access", 10,11,12,14,15,16)]
		event FormOld_QueryEventHandler QueryEvent;

		/// <summary>
		/// SupportByVersion Access 10 11 12 14 15,16
		/// </summary>
		[SupportByVersion("Access", 10,11,12,14,15,16)]
		event FormOld_BeforeQueryEventHandler BeforeQueryEvent;

		/// <summary>
		/// SupportByVersion Access 10 11 12 14 15,16
		/// </summary>
		[SupportByVersion("Access", 10,11,12,14,15,16)]
		event FormOld_SelectionChangeEventHandler SelectionChangeEvent;

		/// <summary>
		/// SupportByVersion Access 10 11 12 14 15,16
		/// </summary>
		[SupportByVersion("Access", 10,11,12,14,15,16)]
		event FormOld_CommandBeforeExecuteEventHandler CommandBeforeExecuteEvent;

		/// <summary>
		/// SupportByVersion Access 10 11 12 14 15,16
		/// </summary>
		[SupportByVersion("Access", 10,11,12,14,15,16)]
		event FormOld_CommandCheckedEventHandler CommandCheckedEvent;

		/// <summary>
		/// SupportByVersion Access 10 11 12 14 15,16
		/// </summary>
		[SupportByVersion("Access", 10,11,12,14,15,16)]
		event FormOld_CommandEnabledEventHandler CommandEnabledEvent;

		/// <summary>
		/// SupportByVersion Access 10 11 12 14 15,16
		/// </summary>
		[SupportByVersion("Access", 10,11,12,14,15,16)]
		event FormOld_CommandExecuteEventHandler CommandExecuteEvent;

		/// <summary>
		/// SupportByVersion Access 10 11 12 14 15,16
		/// </summary>
		[SupportByVersion("Access", 10,11,12,14,15,16)]
		event FormOld_DataSetChangeEventHandler DataSetChangeEvent;

		/// <summary>
		/// SupportByVersion Access 10 11 12 14 15,16
		/// </summary>
		[SupportByVersion("Access", 10,11,12,14,15,16)]
		event FormOld_BeforeScreenTipEventHandler BeforeScreenTipEvent;

		/// <summary>
		/// SupportByVersion Access 10 11 12 14 15,16
		/// </summary>
		[SupportByVersion("Access", 10,11,12,14,15,16)]
		event FormOld_BeforeRenderEventHandler BeforeRenderEvent;

		/// <summary>
		/// SupportByVersion Access 10 11 12 14 15,16
		/// </summary>
		[SupportByVersion("Access", 10,11,12,14,15,16)]
		event FormOld_AfterRenderEventHandler AfterRenderEvent;

		/// <summary>
		/// SupportByVersion Access 10 11 12 14 15,16
		/// </summary>
		[SupportByVersion("Access", 10,11,12,14,15,16)]
		event FormOld_AfterFinalRenderEventHandler AfterFinalRenderEvent;

		/// <summary>
		/// SupportByVersion Access 10 11 12 14 15,16
		/// </summary>
		[SupportByVersion("Access", 10,11,12,14,15,16)]
		event FormOld_AfterLayoutEventHandler AfterLayoutEvent;

		/// <summary>
		/// SupportByVersion Access 10 11 12 14 15,16
		/// </summary>
		[SupportByVersion("Access", 10,11,12,14,15,16)]
		event FormOld_MouseWheelEventHandler MouseWheelEvent;

		/// <summary>
		/// SupportByVersion Access 10 11 12 14 15,16
		/// </summary>
		[SupportByVersion("Access", 10,11,12,14,15,16)]
		event FormOld_ViewChangeEventHandler ViewChangeEvent;

		/// <summary>
		/// SupportByVersion Access 10 11 12 14 15,16
		/// </summary>
		[SupportByVersion("Access", 10,11,12,14,15,16)]
		event FormOld_DataChangeEventHandler DataChangeEvent;

		#endregion
	}
}
