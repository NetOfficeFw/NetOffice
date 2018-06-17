using System;
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
	public delegate void FormOldV10_BeforeCommitTransactionEventHandler(ref Int16 Cancel, ref NetOffice.ADODBApi.Connection connection);
	public delegate void FormOldV10_AfterCommitTransactionEventHandler(ref NetOffice.ADODBApi.Connection connection);
	public delegate void FormOldV10_RollbackTransactionEventHandler(ref NetOffice.ADODBApi.Connection connection);
	public delegate void FormOldV10_OnConnectEventHandler();
	public delegate void FormOldV10_OnDisconnectEventHandler();
	public delegate void FormOldV10_PivotTableChangeEventHandler(Int32 reason);
	public delegate void FormOldV10_QueryEventHandler();
	public delegate void FormOldV10_BeforeQueryEventHandler();
	public delegate void FormOldV10_SelectionChangeEventHandler();
	public delegate void FormOldV10_CommandBeforeExecuteEventHandler(object command, ICOMObject cancel);
	public delegate void FormOldV10_CommandCheckedEventHandler(object command, ICOMObject Checked);
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
    [ComEventContract(typeof(EventContracts._FormEvents))]
	[TypeId("483615A0-74BE-101B-AF4E-00AA003F0F08")]
    public interface FormOldV10 : _Form2, IEventBinding
	{
		#region Events

		/// <summary>
		/// SupportByVersion Access 9 10 11 12 14 15,16
		/// </summary>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		event FormOldV10_LoadEventHandler LoadEvent;

		/// <summary>
		/// SupportByVersion Access 9 10 11 12 14 15,16
		/// </summary>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		event FormOldV10_CurrentEventHandler CurrentEvent;

		/// <summary>
		/// SupportByVersion Access 9 10 11 12 14 15,16
		/// </summary>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		event FormOldV10_BeforeInsertEventHandler BeforeInsertEvent;

		/// <summary>
		/// SupportByVersion Access 9 10 11 12 14 15,16
		/// </summary>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		event FormOldV10_AfterInsertEventHandler AfterInsertEvent;

		/// <summary>
		/// SupportByVersion Access 9 10 11 12 14 15,16
		/// </summary>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		event FormOldV10_BeforeUpdateEventHandler BeforeUpdateEvent;

		/// <summary>
		/// SupportByVersion Access 9 10 11 12 14 15,16
		/// </summary>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		event FormOldV10_AfterUpdateEventHandler AfterUpdateEvent;

		/// <summary>
		/// SupportByVersion Access 9 10 11 12 14 15,16
		/// </summary>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		event FormOldV10_DeleteEventHandler DeleteEvent;

		/// <summary>
		/// SupportByVersion Access 9 10 11 12 14 15,16
		/// </summary>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		event FormOldV10_BeforeDelConfirmEventHandler BeforeDelConfirmEvent;

		/// <summary>
		/// SupportByVersion Access 9 10 11 12 14 15,16
		/// </summary>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		event FormOldV10_AfterDelConfirmEventHandler AfterDelConfirmEvent;

		/// <summary>
		/// SupportByVersion Access 9 10 11 12 14 15,16
		/// </summary>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		event FormOldV10_OpenEventHandler OpenEvent;

		/// <summary>
		/// SupportByVersion Access 9 10 11 12 14 15,16
		/// </summary>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		event FormOldV10_ResizeEventHandler ResizeEvent;

		/// <summary>
		/// SupportByVersion Access 9 10 11 12 14 15,16
		/// </summary>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		event FormOldV10_UnloadEventHandler UnloadEvent;

		/// <summary>
		/// SupportByVersion Access 9 10 11 12 14 15,16
		/// </summary>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		event FormOldV10_CloseEventHandler CloseEvent;

		/// <summary>
		/// SupportByVersion Access 9 10 11 12 14 15,16
		/// </summary>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		event FormOldV10_ActivateEventHandler ActivateEvent;

		/// <summary>
		/// SupportByVersion Access 9 10 11 12 14 15,16
		/// </summary>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		event FormOldV10_DeactivateEventHandler DeactivateEvent;

		/// <summary>
		/// SupportByVersion Access 9 10 11 12 14 15,16
		/// </summary>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		event FormOldV10_GotFocusEventHandler GotFocusEvent;

		/// <summary>
		/// SupportByVersion Access 9 10 11 12 14 15,16
		/// </summary>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		event FormOldV10_LostFocusEventHandler LostFocusEvent;

		/// <summary>
		/// SupportByVersion Access 9 10 11 12 14 15,16
		/// </summary>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		event FormOldV10_ClickEventHandler ClickEvent;

		/// <summary>
		/// SupportByVersion Access 9 10 11 12 14 15,16
		/// </summary>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		event FormOldV10_DblClickEventHandler DblClickEvent;

		/// <summary>
		/// SupportByVersion Access 9 10 11 12 14 15,16
		/// </summary>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		event FormOldV10_MouseDownEventHandler MouseDownEvent;

		/// <summary>
		/// SupportByVersion Access 9 10 11 12 14 15,16
		/// </summary>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		event FormOldV10_MouseMoveEventHandler MouseMoveEvent;

		/// <summary>
		/// SupportByVersion Access 9 10 11 12 14 15,16
		/// </summary>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		event FormOldV10_MouseUpEventHandler MouseUpEvent;

		/// <summary>
		/// SupportByVersion Access 9 10 11 12 14 15,16
		/// </summary>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		event FormOldV10_KeyDownEventHandler KeyDownEvent;

		/// <summary>
		/// SupportByVersion Access 9 10 11 12 14 15,16
		/// </summary>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		event FormOldV10_KeyPressEventHandler KeyPressEvent;

		/// <summary>
		/// SupportByVersion Access 9 10 11 12 14 15,16
		/// </summary>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		event FormOldV10_KeyUpEventHandler KeyUpEvent;

		/// <summary>
		/// SupportByVersion Access 9 10 11 12 14 15,16
		/// </summary>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		event FormOldV10_ErrorEventHandler ErrorEvent;

		/// <summary>
		/// SupportByVersion Access 9 10 11 12 14 15,16
		/// </summary>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		event FormOldV10_TimerEventHandler TimerEvent;

		/// <summary>
		/// SupportByVersion Access 9 10 11 12 14 15,16
		/// </summary>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		event FormOldV10_FilterEventHandler FilterEvent;

		/// <summary>
		/// SupportByVersion Access 9 10 11 12 14 15,16
		/// </summary>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		event FormOldV10_ApplyFilterEventHandler ApplyFilterEvent;

		/// <summary>
		/// SupportByVersion Access 9 10 11 12 14 15,16
		/// </summary>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		event FormOldV10_DirtyEventHandler DirtyEvent;

		/// <summary>
		/// SupportByVersion Access 10 11 12 14 15,16
		/// </summary>
		[SupportByVersion("Access", 10,11,12,14,15,16)]
		event FormOldV10_UndoEventHandler UndoEvent;

		/// <summary>
		/// SupportByVersion Access 10 11 12 14 15,16
		/// </summary>
		[SupportByVersion("Access", 10,11,12,14,15,16)]
		event FormOldV10_RecordExitEventHandler RecordExitEvent;

		/// <summary>
		/// SupportByVersion Access 10 11 12 14 15,16
		/// </summary>
		[SupportByVersion("Access", 10,11,12,14,15,16)]
		event FormOldV10_BeginBatchEditEventHandler BeginBatchEditEvent;

		/// <summary>
		/// SupportByVersion Access 10 11 12 14 15,16
		/// </summary>
		[SupportByVersion("Access", 10,11,12,14,15,16)]
		event FormOldV10_UndoBatchEditEventHandler UndoBatchEditEvent;

		/// <summary>
		/// SupportByVersion Access 10 11 12 14 15,16
		/// </summary>
		[SupportByVersion("Access", 10,11,12,14,15,16)]
		event FormOldV10_BeforeBeginTransactionEventHandler BeforeBeginTransactionEvent;

		/// <summary>
		/// SupportByVersion Access 10 11 12 14 15,16
		/// </summary>
		[SupportByVersion("Access", 10,11,12,14,15,16)]
		event FormOldV10_AfterBeginTransactionEventHandler AfterBeginTransactionEvent;

		/// <summary>
		/// SupportByVersion Access 10 11 12 14 15,16
		/// </summary>
		[SupportByVersion("Access", 10,11,12,14,15,16)]
		event FormOldV10_BeforeCommitTransactionEventHandler BeforeCommitTransactionEvent;

		/// <summary>
		/// SupportByVersion Access 10 11 12 14 15,16
		/// </summary>
		[SupportByVersion("Access", 10,11,12,14,15,16)]
		event FormOldV10_AfterCommitTransactionEventHandler AfterCommitTransactionEvent;

		/// <summary>
		/// SupportByVersion Access 10 11 12 14 15,16
		/// </summary>
		[SupportByVersion("Access", 10,11,12,14,15,16)]
		event FormOldV10_RollbackTransactionEventHandler RollbackTransactionEvent;

		/// <summary>
		/// SupportByVersion Access 10 11 12 14 15,16
		/// </summary>
		[SupportByVersion("Access", 10,11,12,14,15,16)]
		event FormOldV10_OnConnectEventHandler OnConnectEvent;

		/// <summary>
		/// SupportByVersion Access 10 11 12 14 15,16
		/// </summary>
		[SupportByVersion("Access", 10,11,12,14,15,16)]
		event FormOldV10_OnDisconnectEventHandler OnDisconnectEvent;

		/// <summary>
		/// SupportByVersion Access 10 11 12 14 15,16
		/// </summary>
		[SupportByVersion("Access", 10,11,12,14,15,16)]
		event FormOldV10_PivotTableChangeEventHandler PivotTableChangeEvent;

		/// <summary>
		/// SupportByVersion Access 10 11 12 14 15,16
		/// </summary>
		[SupportByVersion("Access", 10,11,12,14,15,16)]
		event FormOldV10_QueryEventHandler QueryEvent;

		/// <summary>
		/// SupportByVersion Access 10 11 12 14 15,16
		/// </summary>
		[SupportByVersion("Access", 10,11,12,14,15,16)]
		event FormOldV10_BeforeQueryEventHandler BeforeQueryEvent;

		/// <summary>
		/// SupportByVersion Access 10 11 12 14 15,16
		/// </summary>
		[SupportByVersion("Access", 10,11,12,14,15,16)]
		event FormOldV10_SelectionChangeEventHandler SelectionChangeEvent;

		/// <summary>
		/// SupportByVersion Access 10 11 12 14 15,16
		/// </summary>
		[SupportByVersion("Access", 10,11,12,14,15,16)]
		event FormOldV10_CommandBeforeExecuteEventHandler CommandBeforeExecuteEvent;

		/// <summary>
		/// SupportByVersion Access 10 11 12 14 15,16
		/// </summary>
		[SupportByVersion("Access", 10,11,12,14,15,16)]
		event FormOldV10_CommandCheckedEventHandler CommandCheckedEvent;

		/// <summary>
		/// SupportByVersion Access 10 11 12 14 15,16
		/// </summary>
		[SupportByVersion("Access", 10,11,12,14,15,16)]
		event FormOldV10_CommandEnabledEventHandler CommandEnabledEvent;

		/// <summary>
		/// SupportByVersion Access 10 11 12 14 15,16
		/// </summary>
		[SupportByVersion("Access", 10,11,12,14,15,16)]
		event FormOldV10_CommandExecuteEventHandler CommandExecuteEvent;

		/// <summary>
		/// SupportByVersion Access 10 11 12 14 15,16
		/// </summary>
		[SupportByVersion("Access", 10,11,12,14,15,16)]
		event FormOldV10_DataSetChangeEventHandler DataSetChangeEvent;

		/// <summary>
		/// SupportByVersion Access 10 11 12 14 15,16
		/// </summary>
		[SupportByVersion("Access", 10,11,12,14,15,16)]
		event FormOldV10_BeforeScreenTipEventHandler BeforeScreenTipEvent;

		/// <summary>
		/// SupportByVersion Access 10 11 12 14 15,16
		/// </summary>
		[SupportByVersion("Access", 10,11,12,14,15,16)]
		event FormOldV10_BeforeRenderEventHandler BeforeRenderEvent;

		/// <summary>
		/// SupportByVersion Access 10 11 12 14 15,16
		/// </summary>
		[SupportByVersion("Access", 10,11,12,14,15,16)]
		event FormOldV10_AfterRenderEventHandler AfterRenderEvent;

		/// <summary>
		/// SupportByVersion Access 10 11 12 14 15,16
		/// </summary>
		[SupportByVersion("Access", 10,11,12,14,15,16)]
		event FormOldV10_AfterFinalRenderEventHandler AfterFinalRenderEvent;

		/// <summary>
		/// SupportByVersion Access 10 11 12 14 15,16
		/// </summary>
		[SupportByVersion("Access", 10,11,12,14,15,16)]
		event FormOldV10_AfterLayoutEventHandler AfterLayoutEvent;

		/// <summary>
		/// SupportByVersion Access 10 11 12 14 15,16
		/// </summary>
		[SupportByVersion("Access", 10,11,12,14,15,16)]
		event FormOldV10_MouseWheelEventHandler MouseWheelEvent;

		/// <summary>
		/// SupportByVersion Access 10 11 12 14 15,16
		/// </summary>
		[SupportByVersion("Access", 10,11,12,14,15,16)]
		event FormOldV10_ViewChangeEventHandler ViewChangeEvent;

		/// <summary>
		/// SupportByVersion Access 10 11 12 14 15,16
		/// </summary>
		[SupportByVersion("Access", 10,11,12,14,15,16)]
		event FormOldV10_DataChangeEventHandler DataChangeEvent;

		#endregion
	}
}
