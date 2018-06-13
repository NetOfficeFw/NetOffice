using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.OutlookApi
{
	#region Delegates

	#pragma warning disable
	public delegate void TaskRequestItem_OpenEventHandler(ref bool cancel);
	public delegate void TaskRequestItem_CustomActionEventHandler(ICOMObject action, ICOMObject response, ref bool cancel);
	public delegate void TaskRequestItem_CustomPropertyChangeEventHandler(string Name);
	public delegate void TaskRequestItem_ForwardEventHandler(ICOMObject forward, ref bool cancel);
	public delegate void TaskRequestItem_CloseEventHandler(ref bool cancel);
	public delegate void TaskRequestItem_PropertyChangeEventHandler(string Name);
	public delegate void TaskRequestItem_ReadEventHandler();
	public delegate void TaskRequestItem_ReplyEventHandler(ICOMObject response, ref bool cancel);
	public delegate void TaskRequestItem_ReplyAllEventHandler(ICOMObject response, ref bool cancel);
	public delegate void TaskRequestItem_SendEventHandler(ref bool cancel);
	public delegate void TaskRequestItem_WriteEventHandler(ref bool cancel);
	public delegate void TaskRequestItem_BeforeCheckNamesEventHandler(ref bool cancel);
	public delegate void TaskRequestItem_AttachmentAddEventHandler(NetOffice.OutlookApi.Attachment attachment);
	public delegate void TaskRequestItem_AttachmentReadEventHandler(NetOffice.OutlookApi.Attachment attachment);
	public delegate void TaskRequestItem_BeforeAttachmentSaveEventHandler(NetOffice.OutlookApi.Attachment attachment, ref bool cancel);
	public delegate void TaskRequestItem_BeforeDeleteEventHandler(ICOMObject item, ref bool cancel);
	public delegate void TaskRequestItem_AttachmentRemoveEventHandler(NetOffice.OutlookApi.Attachment attachment);
	public delegate void TaskRequestItem_BeforeAttachmentAddEventHandler(NetOffice.OutlookApi.Attachment attachment, ref bool cancel);
	public delegate void TaskRequestItem_BeforeAttachmentPreviewEventHandler(NetOffice.OutlookApi.Attachment attachment, ref bool cancel);
	public delegate void TaskRequestItem_BeforeAttachmentReadEventHandler(NetOffice.OutlookApi.Attachment attachment, ref bool cancel);
	public delegate void TaskRequestItem_BeforeAttachmentWriteToTempFileEventHandler(NetOffice.OutlookApi.Attachment attachment, ref bool cancel);
	public delegate void TaskRequestItem_UnloadEventHandler();
	public delegate void TaskRequestItem_BeforeAutoSaveEventHandler(ref bool cancel);
	public delegate void TaskRequestItem_BeforeReadEventHandler();
	public delegate void TaskRequestItem_AfterWriteEventHandler();
	public delegate void TaskRequestItem_ReadCompleteEventHandler(ref bool cancel);
	#pragma warning restore

	#endregion

	/// <summary>
	/// CoClass TaskRequestItem 
	/// SupportByVersion Outlook, 9,10,11,12,14,15,16
	/// </summary>
	/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff862698.aspx </remarks>
	[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
	[EntityType(EntityType.IsCoClass)]
    [ComEventContract(typeof(EventContracts.ItemEvents), typeof(EventContracts.ItemEvents_10))]
	[TypeId("00061050-0000-0000-C000-000000000046")]
    public interface TaskRequestItem : _TaskRequestItem, IEventBinding
	{
        #region Events

        /// <summary>
        /// SupportByVersion Outlook 9 10 11 12 14 15,16
        /// </summary>
        ///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff860373.aspx </remarks>
        [SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		event TaskRequestItem_OpenEventHandler OpenEvent;

		/// <summary>
		/// SupportByVersion Outlook 9 10 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff861909.aspx </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		event TaskRequestItem_CustomActionEventHandler CustomActionEvent;

		/// <summary>
		/// SupportByVersion Outlook 9 10 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff865663.aspx </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		event TaskRequestItem_CustomPropertyChangeEventHandler CustomPropertyChangeEvent;

		/// <summary>
		/// SupportByVersion Outlook 9 10 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff863894.aspx </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		event TaskRequestItem_ForwardEventHandler ForwardEvent;

		/// <summary>
		/// SupportByVersion Outlook 9 10 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff869320.aspx </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		event TaskRequestItem_CloseEventHandler CloseEvent;

		/// <summary>
		/// SupportByVersion Outlook 9 10 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff867884.aspx </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		event TaskRequestItem_PropertyChangeEventHandler PropertyChangeEvent;

		/// <summary>
		/// SupportByVersion Outlook 9 10 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff865285.aspx </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		event TaskRequestItem_ReadEventHandler ReadEvent;

		/// <summary>
		/// SupportByVersion Outlook 9 10 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff868174.aspx </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		event TaskRequestItem_ReplyEventHandler ReplyEvent;

		/// <summary>
		/// SupportByVersion Outlook 9 10 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff869989.aspx </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		event TaskRequestItem_ReplyAllEventHandler ReplyAllEvent;

		/// <summary>
		/// SupportByVersion Outlook 9 10 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff860969.aspx </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		event TaskRequestItem_SendEventHandler SendEvent;

		/// <summary>
		/// SupportByVersion Outlook 9 10 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff869268.aspx </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		event TaskRequestItem_WriteEventHandler WriteEvent;

		/// <summary>
		/// SupportByVersion Outlook 9 10 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff865678.aspx </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		event TaskRequestItem_BeforeCheckNamesEventHandler BeforeCheckNamesEvent;

		/// <summary>
		/// SupportByVersion Outlook 9 10 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff868630.aspx </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		event TaskRequestItem_AttachmentAddEventHandler AttachmentAddEvent;

		/// <summary>
		/// SupportByVersion Outlook 9 10 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff860695.aspx </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		event TaskRequestItem_AttachmentReadEventHandler AttachmentReadEvent;

		/// <summary>
		/// SupportByVersion Outlook 9 10 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff869305.aspx </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		event TaskRequestItem_BeforeAttachmentSaveEventHandler BeforeAttachmentSaveEvent;

		/// <summary>
		/// SupportByVersion Outlook 10 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff868701.aspx </remarks>
		[SupportByVersion("Outlook", 10,11,12,14,15,16)]
		event TaskRequestItem_BeforeDeleteEventHandler BeforeDeleteEvent;

		/// <summary>
		/// SupportByVersion Outlook 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff869723.aspx </remarks>
		[SupportByVersion("Outlook", 12,14,15,16)]
		event TaskRequestItem_AttachmentRemoveEventHandler AttachmentRemoveEvent;

		/// <summary>
		/// SupportByVersion Outlook 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff866413.aspx </remarks>
		[SupportByVersion("Outlook", 12,14,15,16)]
		event TaskRequestItem_BeforeAttachmentAddEventHandler BeforeAttachmentAddEvent;

		/// <summary>
		/// SupportByVersion Outlook 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff863923.aspx </remarks>
		[SupportByVersion("Outlook", 12,14,15,16)]
		event TaskRequestItem_BeforeAttachmentPreviewEventHandler BeforeAttachmentPreviewEvent;

		/// <summary>
		/// SupportByVersion Outlook 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff867589.aspx </remarks>
		[SupportByVersion("Outlook", 12,14,15,16)]
		event TaskRequestItem_BeforeAttachmentReadEventHandler BeforeAttachmentReadEvent;

		/// <summary>
		/// SupportByVersion Outlook 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff869778.aspx </remarks>
		[SupportByVersion("Outlook", 12,14,15,16)]
		event TaskRequestItem_BeforeAttachmentWriteToTempFileEventHandler BeforeAttachmentWriteToTempFileEvent;

		/// <summary>
		/// SupportByVersion Outlook 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff868072.aspx </remarks>
		[SupportByVersion("Outlook", 12,14,15,16)]
		event TaskRequestItem_UnloadEventHandler UnloadEvent;

		/// <summary>
		/// SupportByVersion Outlook 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff860697.aspx </remarks>
		[SupportByVersion("Outlook", 12,14,15,16)]
		event TaskRequestItem_BeforeAutoSaveEventHandler BeforeAutoSaveEvent;

		/// <summary>
		/// SupportByVersion Outlook 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff869784.aspx </remarks>
		[SupportByVersion("Outlook", 14,15,16)]
		event TaskRequestItem_BeforeReadEventHandler BeforeReadEvent;

		/// <summary>
		/// SupportByVersion Outlook 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff867184.aspx </remarks>
		[SupportByVersion("Outlook", 14,15,16)]
		event TaskRequestItem_AfterWriteEventHandler AfterWriteEvent;

		/// <summary>
		/// SupportByVersion Outlook 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/jj228141.aspx </remarks>
		[SupportByVersion("Outlook", 15, 16)]
		event TaskRequestItem_ReadCompleteEventHandler ReadCompleteEvent;

        #endregion
    }
}
