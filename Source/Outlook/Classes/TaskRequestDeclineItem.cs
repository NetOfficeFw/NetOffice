using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.OutlookApi
{
	#region Delegates

	#pragma warning disable
	public delegate void TaskRequestDeclineItem_OpenEventHandler(ref bool cancel);
	public delegate void TaskRequestDeclineItem_CustomActionEventHandler(ICOMObject action, ICOMObject response, ref bool cancel);
	public delegate void TaskRequestDeclineItem_CustomPropertyChangeEventHandler(string name);
	public delegate void TaskRequestDeclineItem_ForwardEventHandler(ICOMObject forward, ref bool cancel);
	public delegate void TaskRequestDeclineItem_CloseEventHandler(ref bool cancel);
	public delegate void TaskRequestDeclineItem_PropertyChangeEventHandler(string Name);
	public delegate void TaskRequestDeclineItem_ReadEventHandler();
	public delegate void TaskRequestDeclineItem_ReplyEventHandler(ICOMObject response, ref bool cancel);
	public delegate void TaskRequestDeclineItem_ReplyAllEventHandler(ICOMObject response, ref bool cancel);
	public delegate void TaskRequestDeclineItem_SendEventHandler(ref bool cancel);
	public delegate void TaskRequestDeclineItem_WriteEventHandler(ref bool cancel);
	public delegate void TaskRequestDeclineItem_BeforeCheckNamesEventHandler(ref bool cancel);
	public delegate void TaskRequestDeclineItem_AttachmentAddEventHandler(NetOffice.OutlookApi.Attachment attachment);
	public delegate void TaskRequestDeclineItem_AttachmentReadEventHandler(NetOffice.OutlookApi.Attachment attachment);
	public delegate void TaskRequestDeclineItem_BeforeAttachmentSaveEventHandler(NetOffice.OutlookApi.Attachment attachment, ref bool cancel);
	public delegate void TaskRequestDeclineItem_BeforeDeleteEventHandler(ICOMObject item, ref bool cancel);
	public delegate void TaskRequestDeclineItem_AttachmentRemoveEventHandler(NetOffice.OutlookApi.Attachment attachment);
	public delegate void TaskRequestDeclineItem_BeforeAttachmentAddEventHandler(NetOffice.OutlookApi.Attachment attachment, ref bool cancel);
	public delegate void TaskRequestDeclineItem_BeforeAttachmentPreviewEventHandler(NetOffice.OutlookApi.Attachment attachment, ref bool cancel);
	public delegate void TaskRequestDeclineItem_BeforeAttachmentReadEventHandler(NetOffice.OutlookApi.Attachment attachment, ref bool cancel);
	public delegate void TaskRequestDeclineItem_BeforeAttachmentWriteToTempFileEventHandler(NetOffice.OutlookApi.Attachment attachment, ref bool cancel);
	public delegate void TaskRequestDeclineItem_UnloadEventHandler();
	public delegate void TaskRequestDeclineItem_BeforeAutoSaveEventHandler(ref bool cancel);
	public delegate void TaskRequestDeclineItem_BeforeReadEventHandler();
	public delegate void TaskRequestDeclineItem_AfterWriteEventHandler();
	public delegate void TaskRequestDeclineItem_ReadCompleteEventHandler(ref bool cancel);
	#pragma warning restore

	#endregion

	/// <summary>
	/// CoClass TaskRequestDeclineItem 
	/// SupportByVersion Outlook, 9,10,11,12,14,15,16
	/// </summary>
	/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff869677.aspx </remarks>
	[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
	[EntityType(EntityType.IsCoClass)]
    [ComEventContract(typeof(EventContracts.ItemEvents), typeof(EventContracts.ItemEvents_10))]
	[TypeId("00061053-0000-0000-C000-000000000046")]
    public interface TaskRequestDeclineItem : _TaskRequestDeclineItem, IEventBinding
	{
        #region Events

        /// <summary>
        /// SupportByVersion Outlook 9 10 11 12 14 15,16
        /// </summary>
        ///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff869953.aspx </remarks>
        [SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		event TaskRequestDeclineItem_OpenEventHandler OpenEvent;

		/// <summary>
		/// SupportByVersion Outlook 9 10 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff868580.aspx </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		event TaskRequestDeclineItem_CustomActionEventHandler CustomActionEvent;

		/// <summary>
		/// SupportByVersion Outlook 9 10 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff867667.aspx </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		event TaskRequestDeclineItem_CustomPropertyChangeEventHandler CustomPropertyChangeEvent;

		/// <summary>
		/// SupportByVersion Outlook 9 10 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff862386.aspx </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		event TaskRequestDeclineItem_ForwardEventHandler ForwardEvent;

		/// <summary>
		/// SupportByVersion Outlook 9 10 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff863613.aspx </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		event TaskRequestDeclineItem_CloseEventHandler CloseEvent;

		/// <summary>
		/// SupportByVersion Outlook 9 10 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff867434.aspx </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		event TaskRequestDeclineItem_PropertyChangeEventHandler PropertyChangeEvent;

		/// <summary>
		/// SupportByVersion Outlook 9 10 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff863455.aspx </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		event TaskRequestDeclineItem_ReadEventHandler ReadEvent;

		/// <summary>
		/// SupportByVersion Outlook 9 10 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff865848.aspx </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		event TaskRequestDeclineItem_ReplyEventHandler ReplyEvent;

		/// <summary>
		/// SupportByVersion Outlook 9 10 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff868811.aspx </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		event TaskRequestDeclineItem_ReplyAllEventHandler ReplyAllEvent;

		/// <summary>
		/// SupportByVersion Outlook 9 10 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff869660.aspx </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		event TaskRequestDeclineItem_SendEventHandler SendEvent;

		/// <summary>
		/// SupportByVersion Outlook 9 10 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff869548.aspx </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		event TaskRequestDeclineItem_WriteEventHandler WriteEvent;

		/// <summary>
		/// SupportByVersion Outlook 9 10 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff869475.aspx </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		event TaskRequestDeclineItem_BeforeCheckNamesEventHandler BeforeCheckNamesEvent;

		/// <summary>
		/// SupportByVersion Outlook 9 10 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff870036.aspx </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		event TaskRequestDeclineItem_AttachmentAddEventHandler AttachmentAddEvent;

		/// <summary>
		/// SupportByVersion Outlook 9 10 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff867870.aspx </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		event TaskRequestDeclineItem_AttachmentReadEventHandler AttachmentReadEvent;

		/// <summary>
		/// SupportByVersion Outlook 9 10 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff860651.aspx </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		event TaskRequestDeclineItem_BeforeAttachmentSaveEventHandler BeforeAttachmentSaveEvent;

		/// <summary>
		/// SupportByVersion Outlook 10 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff868073.aspx </remarks>
		[SupportByVersion("Outlook", 10,11,12,14,15,16)]
		event TaskRequestDeclineItem_BeforeDeleteEventHandler BeforeDeleteEvent;

		/// <summary>
		/// SupportByVersion Outlook 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff866431.aspx </remarks>
		[SupportByVersion("Outlook", 12,14,15,16)]
		event TaskRequestDeclineItem_AttachmentRemoveEventHandler AttachmentRemoveEvent;

		/// <summary>
		/// SupportByVersion Outlook 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff861832.aspx </remarks>
		[SupportByVersion("Outlook", 12,14,15,16)]
		event TaskRequestDeclineItem_BeforeAttachmentAddEventHandler BeforeAttachmentAddEvent;

		/// <summary>
		/// SupportByVersion Outlook 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff863320.aspx </remarks>
		[SupportByVersion("Outlook", 12,14,15,16)]
		event TaskRequestDeclineItem_BeforeAttachmentPreviewEventHandler BeforeAttachmentPreviewEvent;

		/// <summary>
		/// SupportByVersion Outlook 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff869692.aspx </remarks>
		[SupportByVersion("Outlook", 12,14,15,16)]
		event TaskRequestDeclineItem_BeforeAttachmentReadEventHandler BeforeAttachmentReadEvent;

		/// <summary>
		/// SupportByVersion Outlook 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff869074.aspx </remarks>
		[SupportByVersion("Outlook", 12,14,15,16)]
		event TaskRequestDeclineItem_BeforeAttachmentWriteToTempFileEventHandler BeforeAttachmentWriteToTempFileEvent;

		/// <summary>
		/// SupportByVersion Outlook 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff862961.aspx </remarks>
		[SupportByVersion("Outlook", 12,14,15,16)]
		event TaskRequestDeclineItem_UnloadEventHandler UnloadEvent;

		/// <summary>
		/// SupportByVersion Outlook 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff868271.aspx </remarks>
		[SupportByVersion("Outlook", 12,14,15,16)]
		event TaskRequestDeclineItem_BeforeAutoSaveEventHandler BeforeAutoSaveEvent;

		/// <summary>
		/// SupportByVersion Outlook 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff867893.aspx </remarks>
		[SupportByVersion("Outlook", 14,15,16)]
		event TaskRequestDeclineItem_BeforeReadEventHandler BeforeReadEvent;

		/// <summary>
		/// SupportByVersion Outlook 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff861041.aspx </remarks>
		[SupportByVersion("Outlook", 14,15,16)]
		event TaskRequestDeclineItem_AfterWriteEventHandler AfterWriteEvent;

		/// <summary>
		/// SupportByVersion Outlook 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/jj229667.aspx </remarks>
		[SupportByVersion("Outlook", 15, 16)]
		event TaskRequestDeclineItem_ReadCompleteEventHandler ReadCompleteEvent;

        #endregion
    }
}
