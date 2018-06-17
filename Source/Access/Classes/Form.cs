using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.AccessApi
{
	#region Delegates

	#pragma warning disable
	public delegate void Form_LoadEventHandler();
	public delegate void Form_CurrentEventHandler();
	public delegate void Form_BeforeInsertEventHandler(ref Int16 cancel);
	public delegate void Form_AfterInsertEventHandler();
	public delegate void Form_BeforeUpdateEventHandler(ref Int16 cancel);
	public delegate void Form_AfterUpdateEventHandler();
	public delegate void Form_DeleteEventHandler(ref Int16 cancel);
	public delegate void Form_BeforeDelConfirmEventHandler(ref Int16 cancel, ref Int16 response);
	public delegate void Form_AfterDelConfirmEventHandler(ref Int16 status);
	public delegate void Form_OpenEventHandler(ref Int16 cancel);
	public delegate void Form_ResizeEventHandler();
	public delegate void Form_UnloadEventHandler(ref Int16 cancel);
	public delegate void Form_CloseEventHandler();
	public delegate void Form_ActivateEventHandler();
	public delegate void Form_DeactivateEventHandler();
	public delegate void Form_GotFocusEventHandler();
	public delegate void Form_LostFocusEventHandler();
	public delegate void Form_ClickEventHandler();
	public delegate void Form_DblClickEventHandler(ref Int16 cancel);
	public delegate void Form_MouseDownEventHandler(ref Int16 button, ref Int16 shift, ref Single x, ref Single y);
	public delegate void Form_MouseMoveEventHandler(ref Int16 button, ref Int16 shift, ref Single x, ref Single y);
	public delegate void Form_MouseUpEventHandler(ref Int16 button, ref Int16 shift, ref Single x, ref Single y);
	public delegate void Form_KeyDownEventHandler(ref Int16 keyCode, ref Int16 shift);
	public delegate void Form_KeyPressEventHandler(ref Int16 KeyAscii);
	public delegate void Form_KeyUpEventHandler(ref Int16 keyCode, ref Int16 shift);
	public delegate void Form_ErrorEventHandler(ref Int16 dataErr, ref Int16 response);
	public delegate void Form_TimerEventHandler();
	public delegate void Form_FilterEventHandler(ref Int16 cancel, ref Int16 filterType);
	public delegate void Form_ApplyFilterEventHandler(ref Int16 cancel, ref Int16 applyType);
	public delegate void Form_DirtyEventHandler(ref Int16 cancel);
	public delegate void Form_UndoEventHandler(ref Int16 cancel);
	public delegate void Form_RecordExitEventHandler(ref Int16 cancel);
	public delegate void Form_BeginBatchEditEventHandler(ref Int16 cancel);
	public delegate void Form_UndoBatchEditEventHandler(ref Int16 cancel);
	public delegate void Form_BeforeBeginTransactionEventHandler(ref Int16 cancel, ref NetOffice.ADODBApi.Connection connection);
	public delegate void Form_AfterBeginTransactionEventHandler(ref NetOffice.ADODBApi.Connection connection);
	public delegate void Form_BeforeCommitTransactionEventHandler(ref Int16 Cancel, ref NetOffice.ADODBApi.Connection connection);
	public delegate void Form_AfterCommitTransactionEventHandler(ref NetOffice.ADODBApi.Connection connection);
	public delegate void Form_RollbackTransactionEventHandler(ref NetOffice.ADODBApi.Connection connection);
	public delegate void Form_OnConnectEventHandler();
	public delegate void Form_OnDisconnectEventHandler();
	public delegate void Form_PivotTableChangeEventHandler(Int32 reason);
	public delegate void Form_QueryEventHandler();
	public delegate void Form_BeforeQueryEventHandler();
	public delegate void Form_SelectionChangeEventHandler();
	public delegate void Form_CommandBeforeExecuteEventHandler(object command, ICOMObject cancel);
	public delegate void Form_CommandCheckedEventHandler(object command, ICOMObject Checked);
	public delegate void Form_CommandEnabledEventHandler(object command, ICOMObject Enabled);
	public delegate void Form_CommandExecuteEventHandler(object command);
	public delegate void Form_DataSetChangeEventHandler();
	public delegate void Form_BeforeScreenTipEventHandler(ICOMObject screenTipText, ICOMObject sourceObject);
	public delegate void Form_BeforeRenderEventHandler(ICOMObject drawObject, ICOMObject chartObject, ICOMObject cancel);
	public delegate void Form_AfterRenderEventHandler(ICOMObject drawObject, ICOMObject chartObject);
	public delegate void Form_AfterFinalRenderEventHandler(ICOMObject drawObject);
	public delegate void Form_AfterLayoutEventHandler(ICOMObject drawObject);
	public delegate void Form_MouseWheelEventHandler(bool page, Int32 count);
	public delegate void Form_ViewChangeEventHandler(Int32 reason);
	public delegate void Form_DataChangeEventHandler(Int32 reason);
	#pragma warning restore

	#endregion

	/// <summary>
	/// CoClass Form 
	/// SupportByVersion Access, 9,10,11,12,14,15,16
	/// </summary>
	/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff195841.aspx </remarks>
	[SupportByVersion("Access", 9,10,11,12,14,15,16)]
	[EntityType(EntityType.IsCoClass)]
    [ComEventContract(typeof(EventContracts._FormEvents), typeof(EventContracts._FormEvents2))]
	[TypeId("483615A0-74BE-101B-AF4E-00AA003F0F07")]
    public interface Form : _Form3, IEventBinding
	{
		#region Events

		/// <summary>
		/// SupportByVersion Access 9 10 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff821347.aspx </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		event Form_LoadEventHandler LoadEvent;

		/// <summary>
		/// SupportByVersion Access 9 10 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff193159.aspx </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		event Form_CurrentEventHandler CurrentEvent;

		/// <summary>
		/// SupportByVersion Access 9 10 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff835397.aspx </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		event Form_BeforeInsertEventHandler BeforeInsertEvent;

		/// <summary>
		/// SupportByVersion Access 9 10 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff844951.aspx </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		event Form_AfterInsertEventHandler AfterInsertEvent;

		/// <summary>
		/// SupportByVersion Access 9 10 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff822421.aspx </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		event Form_BeforeUpdateEventHandler BeforeUpdateEvent;

		/// <summary>
		/// SupportByVersion Access 9 10 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff822097.aspx </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		event Form_AfterUpdateEventHandler AfterUpdateEvent;

		/// <summary>
		/// SupportByVersion Access 9 10 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff197077.aspx </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		event Form_DeleteEventHandler DeleteEvent;

		/// <summary>
		/// SupportByVersion Access 9 10 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff192478.aspx </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		event Form_BeforeDelConfirmEventHandler BeforeDelConfirmEvent;

		/// <summary>
		/// SupportByVersion Access 9 10 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff193473.aspx </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		event Form_AfterDelConfirmEventHandler AfterDelConfirmEvent;

		/// <summary>
		/// SupportByVersion Access 9 10 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff196808.aspx </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		event Form_OpenEventHandler OpenEvent;

		/// <summary>
		/// SupportByVersion Access 9 10 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff835413.aspx </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		event Form_ResizeEventHandler ResizeEvent;

		/// <summary>
		/// SupportByVersion Access 9 10 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff845441.aspx </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		event Form_UnloadEventHandler UnloadEvent;

		/// <summary>
		/// SupportByVersion Access 9 10 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff835965.aspx </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		event Form_CloseEventHandler CloseEvent;

		/// <summary>
		/// SupportByVersion Access 9 10 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff845446.aspx </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		event Form_ActivateEventHandler ActivateEvent;

		/// <summary>
		/// SupportByVersion Access 9 10 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff197024.aspx </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		event Form_DeactivateEventHandler DeactivateEvent;

		/// <summary>
		/// SupportByVersion Access 9 10 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff835433.aspx </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		event Form_GotFocusEventHandler GotFocusEvent;

		/// <summary>
		/// SupportByVersion Access 9 10 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff845072.aspx </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		event Form_LostFocusEventHandler LostFocusEvent;

		/// <summary>
		/// SupportByVersion Access 9 10 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff193141.aspx </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		event Form_ClickEventHandler ClickEvent;

		/// <summary>
		/// SupportByVersion Access 9 10 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff822517.aspx </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		event Form_DblClickEventHandler DblClickEvent;

		/// <summary>
		/// SupportByVersion Access 9 10 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff845076.aspx </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		event Form_MouseDownEventHandler MouseDownEvent;

		/// <summary>
		/// SupportByVersion Access 9 10 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff835703.aspx </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		event Form_MouseMoveEventHandler MouseMoveEvent;

		/// <summary>
		/// SupportByVersion Access 9 10 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff822046.aspx </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		event Form_MouseUpEventHandler MouseUpEvent;

		/// <summary>
		/// SupportByVersion Access 9 10 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff834507.aspx </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		event Form_KeyDownEventHandler KeyDownEvent;

		/// <summary>
		/// SupportByVersion Access 9 10 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff194937.aspx </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		event Form_KeyPressEventHandler KeyPressEvent;

		/// <summary>
		/// SupportByVersion Access 9 10 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff835350.aspx </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		event Form_KeyUpEventHandler KeyUpEvent;

		/// <summary>
		/// SupportByVersion Access 9 10 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff836345.aspx </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		event Form_ErrorEventHandler ErrorEvent;

		/// <summary>
		/// SupportByVersion Access 9 10 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff192530.aspx </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		event Form_TimerEventHandler TimerEvent;

		/// <summary>
		/// SupportByVersion Access 9 10 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff191892.aspx </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		event Form_FilterEventHandler FilterEvent;

		/// <summary>
		/// SupportByVersion Access 9 10 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff823191.aspx </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		event Form_ApplyFilterEventHandler ApplyFilterEvent;

		/// <summary>
		/// SupportByVersion Access 9 10 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff835655.aspx </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		event Form_DirtyEventHandler DirtyEvent;

		/// <summary>
		/// SupportByVersion Access 10 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff837253.aspx </remarks>
		[SupportByVersion("Access", 10,11,12,14,15,16)]
		event Form_UndoEventHandler UndoEvent;

		/// <summary>
		/// SupportByVersion Access 10 11 12 14 15,16
		/// </summary>
		[SupportByVersion("Access", 10,11,12,14,15,16)]
		event Form_RecordExitEventHandler RecordExitEvent;

		/// <summary>
		/// SupportByVersion Access 10 11 12 14 15,16
		/// </summary>
		[SupportByVersion("Access", 10,11,12,14,15,16)]
		event Form_BeginBatchEditEventHandler BeginBatchEditEvent;

		/// <summary>
		/// SupportByVersion Access 10 11 12 14 15,16
		/// </summary>
		[SupportByVersion("Access", 10,11,12,14,15,16)]
		event Form_UndoBatchEditEventHandler UndoBatchEditEvent;

		/// <summary>
		/// SupportByVersion Access 10 11 12 14 15,16
		/// </summary>
		[SupportByVersion("Access", 10,11,12,14,15,16)]
		event Form_BeforeBeginTransactionEventHandler BeforeBeginTransactionEvent;

		/// <summary>
		/// SupportByVersion Access 10 11 12 14 15,16
		/// </summary>
		[SupportByVersion("Access", 10,11,12,14,15,16)]
		event Form_AfterBeginTransactionEventHandler AfterBeginTransactionEvent;

		/// <summary>
		/// SupportByVersion Access 10 11 12 14 15,16
		/// </summary>
		[SupportByVersion("Access", 10,11,12,14,15,16)]
		event Form_BeforeCommitTransactionEventHandler BeforeCommitTransactionEvent;

		/// <summary>
		/// SupportByVersion Access 10 11 12 14 15,16
		/// </summary>
		[SupportByVersion("Access", 10,11,12,14,15,16)]
		event Form_AfterCommitTransactionEventHandler AfterCommitTransactionEvent;

		/// <summary>
		/// SupportByVersion Access 10 11 12 14 15,16
		/// </summary>
		[SupportByVersion("Access", 10,11,12,14,15,16)]
		event Form_RollbackTransactionEventHandler RollbackTransactionEvent;

		/// <summary>
		/// SupportByVersion Access 10 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff192637.aspx </remarks>
		[SupportByVersion("Access", 10,11,12,14,15,16)]
		event Form_OnConnectEventHandler OnConnectEvent;

		/// <summary>
		/// SupportByVersion Access 10 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff822088.aspx </remarks>
		[SupportByVersion("Access", 10,11,12,14,15,16)]
		event Form_OnDisconnectEventHandler OnDisconnectEvent;

		/// <summary>
		/// SupportByVersion Access 10 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff197316.aspx </remarks>
		[SupportByVersion("Access", 10,11,12,14,15,16)]
		event Form_PivotTableChangeEventHandler PivotTableChangeEvent;

		/// <summary>
		/// SupportByVersion Access 10 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff836647.aspx </remarks>
		[SupportByVersion("Access", 10,11,12,14,15,16)]
		event Form_QueryEventHandler QueryEvent;

		/// <summary>
		/// SupportByVersion Access 10 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff844995.aspx </remarks>
		[SupportByVersion("Access", 10,11,12,14,15,16)]
		event Form_BeforeQueryEventHandler BeforeQueryEvent;

		/// <summary>
		/// SupportByVersion Access 10 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff193548.aspx </remarks>
		[SupportByVersion("Access", 10,11,12,14,15,16)]
		event Form_SelectionChangeEventHandler SelectionChangeEvent;

		/// <summary>
		/// SupportByVersion Access 10 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff193789.aspx </remarks>
		[SupportByVersion("Access", 10,11,12,14,15,16)]
		event Form_CommandBeforeExecuteEventHandler CommandBeforeExecuteEvent;

		/// <summary>
		/// SupportByVersion Access 10 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff836292.aspx </remarks>
		[SupportByVersion("Access", 10,11,12,14,15,16)]
		event Form_CommandCheckedEventHandler CommandCheckedEvent;

		/// <summary>
		/// SupportByVersion Access 10 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff193487.aspx </remarks>
		[SupportByVersion("Access", 10,11,12,14,15,16)]
		event Form_CommandEnabledEventHandler CommandEnabledEvent;

		/// <summary>
		/// SupportByVersion Access 10 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff822074.aspx </remarks>
		[SupportByVersion("Access", 10,11,12,14,15,16)]
		event Form_CommandExecuteEventHandler CommandExecuteEvent;

		/// <summary>
		/// SupportByVersion Access 10 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff822022.aspx </remarks>
		[SupportByVersion("Access", 10,11,12,14,15,16)]
		event Form_DataSetChangeEventHandler DataSetChangeEvent;

		/// <summary>
		/// SupportByVersion Access 10 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff845031.aspx </remarks>
		[SupportByVersion("Access", 10,11,12,14,15,16)]
		event Form_BeforeScreenTipEventHandler BeforeScreenTipEvent;

		/// <summary>
		/// SupportByVersion Access 10 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff194217.aspx </remarks>
		[SupportByVersion("Access", 10,11,12,14,15,16)]
		event Form_BeforeRenderEventHandler BeforeRenderEvent;

		/// <summary>
		/// SupportByVersion Access 10 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff192300.aspx </remarks>
		[SupportByVersion("Access", 10,11,12,14,15,16)]
		event Form_AfterRenderEventHandler AfterRenderEvent;

		/// <summary>
		/// SupportByVersion Access 10 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff197090.aspx </remarks>
		[SupportByVersion("Access", 10,11,12,14,15,16)]
		event Form_AfterFinalRenderEventHandler AfterFinalRenderEvent;

		/// <summary>
		/// SupportByVersion Access 10 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff192662.aspx </remarks>
		[SupportByVersion("Access", 10,11,12,14,15,16)]
		event Form_AfterLayoutEventHandler AfterLayoutEvent;

		/// <summary>
		/// SupportByVersion Access 10 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff836387.aspx </remarks>
		[SupportByVersion("Access", 10,11,12,14,15,16)]
		event Form_MouseWheelEventHandler MouseWheelEvent;

		/// <summary>
		/// SupportByVersion Access 10 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff821041.aspx </remarks>
		[SupportByVersion("Access", 10,11,12,14,15,16)]
		event Form_ViewChangeEventHandler ViewChangeEvent;

		/// <summary>
		/// SupportByVersion Access 10 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff844766.aspx </remarks>
		[SupportByVersion("Access", 10,11,12,14,15,16)]
		event Form_DataChangeEventHandler DataChangeEvent;

		#endregion
	}
}
