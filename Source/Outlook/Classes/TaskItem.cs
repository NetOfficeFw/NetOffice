using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.OutlookApi
{
	#region Delegates

	#pragma warning disable
	public delegate void TaskItem_OpenEventHandler(ref bool cancel);
	public delegate void TaskItem_CustomActionEventHandler(ICOMObject action, ICOMObject response, ref bool cancel);
	public delegate void TaskItem_CustomPropertyChangeEventHandler(string name);
	public delegate void TaskItem_ForwardEventHandler(ICOMObject forward, ref bool cancel);
	public delegate void TaskItem_CloseEventHandler(ref bool cancel);
	public delegate void TaskItem_PropertyChangeEventHandler(string Name);
	public delegate void TaskItem_ReadEventHandler();
	public delegate void TaskItem_ReplyEventHandler(ICOMObject response, ref bool cancel);
	public delegate void TaskItem_ReplyAllEventHandler(ICOMObject response, ref bool cancel);
	public delegate void TaskItem_SendEventHandler(ref bool cancel);
	public delegate void TaskItem_WriteEventHandler(ref bool cancel);
	public delegate void TaskItem_BeforeCheckNamesEventHandler(ref bool cancel);
	public delegate void TaskItem_AttachmentAddEventHandler(NetOffice.OutlookApi.Attachment attachment);
	public delegate void TaskItem_AttachmentReadEventHandler(NetOffice.OutlookApi.Attachment attachment);
	public delegate void TaskItem_BeforeAttachmentSaveEventHandler(NetOffice.OutlookApi.Attachment attachment, ref bool cancel);
	public delegate void TaskItem_BeforeDeleteEventHandler(ICOMObject item, ref bool cancel);
	public delegate void TaskItem_AttachmentRemoveEventHandler(NetOffice.OutlookApi.Attachment attachment);
	public delegate void TaskItem_BeforeAttachmentAddEventHandler(NetOffice.OutlookApi.Attachment attachment, ref bool cancel);
	public delegate void TaskItem_BeforeAttachmentPreviewEventHandler(NetOffice.OutlookApi.Attachment attachment, ref bool cancel);
	public delegate void TaskItem_BeforeAttachmentReadEventHandler(NetOffice.OutlookApi.Attachment attachment, ref bool cancel);
	public delegate void TaskItem_BeforeAttachmentWriteToTempFileEventHandler(NetOffice.OutlookApi.Attachment attachment, ref bool cancel);
	public delegate void TaskItem_UnloadEventHandler();
	public delegate void TaskItem_BeforeAutoSaveEventHandler(ref bool cancel);
	public delegate void TaskItem_BeforeReadEventHandler();
	public delegate void TaskItem_AfterWriteEventHandler();
	public delegate void TaskItem_ReadCompleteEventHandler(ref bool cancel);
	#pragma warning restore

	#endregion

	/// <summary>
	/// CoClass TaskItem 
	/// SupportByVersion Outlook, 9,10,11,12,14,15,16
	/// </summary>
	/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff865624.aspx </remarks>
	[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
	[EntityType(EntityType.IsCoClass)]
    [ComEventContract(typeof(EventContracts.ItemEvents), typeof(EventContracts.ItemEvents_10))]
	[TypeId("00061032-0000-0000-C000-000000000046")]
    public interface TaskItem : _TaskItem, IEventBinding
	{
        #region Events

        /// <summary>
        /// SupportByVersion Outlook 9 10 11 12 14 15,16
        /// </summary>
        ///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff860296.aspx </remarks>
        [SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		event TaskItem_OpenEventHandler OpenEvent;

		/// <summary>
		/// SupportByVersion Outlook 9 10 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff866242.aspx </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		event TaskItem_CustomActionEventHandler CustomActionEvent;

		/// <summary>
		/// SupportByVersion Outlook 9 10 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff868674.aspx </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		event TaskItem_CustomPropertyChangeEventHandler CustomPropertyChangeEvent;

		/// <summary>
		/// SupportByVersion Outlook 9 10 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff867825.aspx </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		event TaskItem_ForwardEventHandler ForwardEvent;

		/// <summary>
		/// SupportByVersion Outlook 9 10 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff868282.aspx </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		event TaskItem_CloseEventHandler CloseEvent;

		/// <summary>
		/// SupportByVersion Outlook 9 10 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff868530.aspx </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		event TaskItem_PropertyChangeEventHandler PropertyChangeEvent;

		/// <summary>
		/// SupportByVersion Outlook 9 10 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff867392.aspx </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		event TaskItem_ReadEventHandler ReadEvent;

		/// <summary>
		/// SupportByVersion Outlook 9 10 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff865643.aspx </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		event TaskItem_ReplyEventHandler ReplyEvent;

		/// <summary>
		/// SupportByVersion Outlook 9 10 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff870159.aspx </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		event TaskItem_ReplyAllEventHandler ReplyAllEvent;

		/// <summary>
		/// SupportByVersion Outlook 9 10 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff869978.aspx </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		event TaskItem_SendEventHandler SendEvent;

		/// <summary>
		/// SupportByVersion Outlook 9 10 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff862720.aspx </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		event TaskItem_WriteEventHandler WriteEvent;

		/// <summary>
		/// SupportByVersion Outlook 9 10 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff868404.aspx </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		event TaskItem_BeforeCheckNamesEventHandler BeforeCheckNamesEvent;

		/// <summary>
		/// SupportByVersion Outlook 9 10 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff868022.aspx </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		event TaskItem_AttachmentAddEventHandler AttachmentAddEvent;

		/// <summary>
		/// SupportByVersion Outlook 9 10 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff867439.aspx </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		event TaskItem_AttachmentReadEventHandler AttachmentReadEvent;

		/// <summary>
		/// SupportByVersion Outlook 9 10 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff867830.aspx </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		event TaskItem_BeforeAttachmentSaveEventHandler BeforeAttachmentSaveEvent;

		/// <summary>
		/// SupportByVersion Outlook 10 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff868852.aspx </remarks>
		[SupportByVersion("Outlook", 10,11,12,14,15,16)]
		event TaskItem_BeforeDeleteEventHandler BeforeDeleteEvent;

		/// <summary>
		/// SupportByVersion Outlook 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff862707.aspx </remarks>
		[SupportByVersion("Outlook", 12,14,15,16)]
		event TaskItem_AttachmentRemoveEventHandler AttachmentRemoveEvent;

		/// <summary>
		/// SupportByVersion Outlook 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff869511.aspx </remarks>
		[SupportByVersion("Outlook", 12,14,15,16)]
		event TaskItem_BeforeAttachmentAddEventHandler BeforeAttachmentAddEvent;

		/// <summary>
		/// SupportByVersion Outlook 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff865650.aspx </remarks>
		[SupportByVersion("Outlook", 12,14,15,16)]
		event TaskItem_BeforeAttachmentPreviewEventHandler BeforeAttachmentPreviewEvent;

		/// <summary>
		/// SupportByVersion Outlook 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff862710.aspx </remarks>
		[SupportByVersion("Outlook", 12,14,15,16)]
		event TaskItem_BeforeAttachmentReadEventHandler BeforeAttachmentReadEvent;

		/// <summary>
		/// SupportByVersion Outlook 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff866393.aspx </remarks>
		[SupportByVersion("Outlook", 12,14,15,16)]
		event TaskItem_BeforeAttachmentWriteToTempFileEventHandler BeforeAttachmentWriteToTempFileEvent;

		/// <summary>
		/// SupportByVersion Outlook 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff870190.aspx </remarks>
		[SupportByVersion("Outlook", 12,14,15,16)]
		event TaskItem_UnloadEventHandler UnloadEvent;

		/// <summary>
		/// SupportByVersion Outlook 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff863617.aspx </remarks>
		[SupportByVersion("Outlook", 12,14,15,16)]
		event TaskItem_BeforeAutoSaveEventHandler BeforeAutoSaveEvent;

		/// <summary>
		/// SupportByVersion Outlook 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff868569.aspx </remarks>
		[SupportByVersion("Outlook", 14,15,16)]
		event TaskItem_BeforeReadEventHandler BeforeReadEvent;

		/// <summary>
		/// SupportByVersion Outlook 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff868184.aspx </remarks>
		[SupportByVersion("Outlook", 14,15,16)]
		event TaskItem_AfterWriteEventHandler AfterWriteEvent;

		/// <summary>
		/// SupportByVersion Outlook 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/jj227312.aspx </remarks>
		[SupportByVersion("Outlook", 15, 16)]
		event TaskItem_ReadCompleteEventHandler ReadCompleteEvent;

        #endregion
    }
}
