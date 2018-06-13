using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.OutlookApi
{
	#region Delegates

	#pragma warning disable
	public delegate void TaskRequestUpdateItem_OpenEventHandler(ref bool cancel);
	public delegate void TaskRequestUpdateItem_CustomActionEventHandler(ICOMObject action, ICOMObject response, ref bool cancel);
	public delegate void TaskRequestUpdateItem_CustomPropertyChangeEventHandler(string name);
	public delegate void TaskRequestUpdateItem_ForwardEventHandler(ICOMObject forward, ref bool cancel);
	public delegate void TaskRequestUpdateItem_CloseEventHandler(ref bool cancel);
	public delegate void TaskRequestUpdateItem_PropertyChangeEventHandler(string name);
	public delegate void TaskRequestUpdateItem_ReadEventHandler();
	public delegate void TaskRequestUpdateItem_ReplyEventHandler(ICOMObject response, ref bool cancel);
	public delegate void TaskRequestUpdateItem_ReplyAllEventHandler(ICOMObject response, ref bool cancel);
	public delegate void TaskRequestUpdateItem_SendEventHandler(ref bool cancel);
	public delegate void TaskRequestUpdateItem_WriteEventHandler(ref bool cancel);
	public delegate void TaskRequestUpdateItem_BeforeCheckNamesEventHandler(ref bool cancel);
	public delegate void TaskRequestUpdateItem_AttachmentAddEventHandler(NetOffice.OutlookApi.Attachment attachment);
	public delegate void TaskRequestUpdateItem_AttachmentReadEventHandler(NetOffice.OutlookApi.Attachment attachment);
	public delegate void TaskRequestUpdateItem_BeforeAttachmentSaveEventHandler(NetOffice.OutlookApi.Attachment attachment, ref bool cancel);
	public delegate void TaskRequestUpdateItem_BeforeDeleteEventHandler(ICOMObject item, ref bool cancel);
	public delegate void TaskRequestUpdateItem_AttachmentRemoveEventHandler(NetOffice.OutlookApi.Attachment attachment);
	public delegate void TaskRequestUpdateItem_BeforeAttachmentAddEventHandler(NetOffice.OutlookApi.Attachment attachment, ref bool cancel);
	public delegate void TaskRequestUpdateItem_BeforeAttachmentPreviewEventHandler(NetOffice.OutlookApi.Attachment attachment, ref bool cancel);
	public delegate void TaskRequestUpdateItem_BeforeAttachmentReadEventHandler(NetOffice.OutlookApi.Attachment attachment, ref bool cancel);
	public delegate void TaskRequestUpdateItem_BeforeAttachmentWriteToTempFileEventHandler(NetOffice.OutlookApi.Attachment attachment, ref bool cancel);
	public delegate void TaskRequestUpdateItem_UnloadEventHandler();
	public delegate void TaskRequestUpdateItem_BeforeAutoSaveEventHandler(ref bool cancel);
	public delegate void TaskRequestUpdateItem_BeforeReadEventHandler();
	public delegate void TaskRequestUpdateItem_AfterWriteEventHandler();
	public delegate void TaskRequestUpdateItem_ReadCompleteEventHandler(ref bool cancel);
	#pragma warning restore

	#endregion

	/// <summary>
	/// CoClass TaskRequestUpdateItem 
	/// SupportByVersion Outlook, 9,10,11,12,14,15,16
	/// </summary>
	/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff865401.aspx </remarks>
	[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
	[EntityType(EntityType.IsCoClass)]
    [ComEventContract(typeof(EventContracts.ItemEvents), typeof(EventContracts.ItemEvents_10))]
	[TypeId("00061051-0000-0000-C000-000000000046")]
    public interface TaskRequestUpdateItem : _TaskRequestUpdateItem, IEventBinding
	{
        #region Events

        /// <summary>
        /// SupportByVersion Outlook 9 10 11 12 14 15,16
        /// </summary>
        ///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff866192.aspx </remarks>
        [SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		event TaskRequestUpdateItem_OpenEventHandler OpenEvent;

		/// <summary>
		/// SupportByVersion Outlook 9 10 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff867640.aspx </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		event TaskRequestUpdateItem_CustomActionEventHandler CustomActionEvent;

		/// <summary>
		/// SupportByVersion Outlook 9 10 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff863309.aspx </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		event TaskRequestUpdateItem_CustomPropertyChangeEventHandler CustomPropertyChangeEvent;

		/// <summary>
		/// SupportByVersion Outlook 9 10 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff869078.aspx </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		event TaskRequestUpdateItem_ForwardEventHandler ForwardEvent;

		/// <summary>
		/// SupportByVersion Outlook 9 10 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff868010.aspx </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		event TaskRequestUpdateItem_CloseEventHandler CloseEvent;

		/// <summary>
		/// SupportByVersion Outlook 9 10 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff864400.aspx </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		event TaskRequestUpdateItem_PropertyChangeEventHandler PropertyChangeEvent;

		/// <summary>
		/// SupportByVersion Outlook 9 10 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff869913.aspx </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		event TaskRequestUpdateItem_ReadEventHandler ReadEvent;

		/// <summary>
		/// SupportByVersion Outlook 9 10 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff868703.aspx </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		event TaskRequestUpdateItem_ReplyEventHandler ReplyEvent;

		/// <summary>
		/// SupportByVersion Outlook 9 10 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff868680.aspx </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		event TaskRequestUpdateItem_ReplyAllEventHandler ReplyAllEvent;

		/// <summary>
		/// SupportByVersion Outlook 9 10 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff865381.aspx </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		event TaskRequestUpdateItem_SendEventHandler SendEvent;

		/// <summary>
		/// SupportByVersion Outlook 9 10 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff868563.aspx </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		event TaskRequestUpdateItem_WriteEventHandler WriteEvent;

		/// <summary>
		/// SupportByVersion Outlook 9 10 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff869584.aspx </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		event TaskRequestUpdateItem_BeforeCheckNamesEventHandler BeforeCheckNamesEvent;

		/// <summary>
		/// SupportByVersion Outlook 9 10 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff866907.aspx </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		event TaskRequestUpdateItem_AttachmentAddEventHandler AttachmentAddEvent;

		/// <summary>
		/// SupportByVersion Outlook 9 10 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff865389.aspx </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		event TaskRequestUpdateItem_AttachmentReadEventHandler AttachmentReadEvent;

		/// <summary>
		/// SupportByVersion Outlook 9 10 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff862977.aspx </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		event TaskRequestUpdateItem_BeforeAttachmentSaveEventHandler BeforeAttachmentSaveEvent;

		/// <summary>
		/// SupportByVersion Outlook 10 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff868597.aspx </remarks>
		[SupportByVersion("Outlook", 10,11,12,14,15,16)]
		event TaskRequestUpdateItem_BeforeDeleteEventHandler BeforeDeleteEvent;

		/// <summary>
		/// SupportByVersion Outlook 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff867670.aspx </remarks>
		[SupportByVersion("Outlook", 12,14,15,16)]
		event TaskRequestUpdateItem_AttachmentRemoveEventHandler AttachmentRemoveEvent;

		/// <summary>
		/// SupportByVersion Outlook 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff868619.aspx </remarks>
		[SupportByVersion("Outlook", 12,14,15,16)]
		event TaskRequestUpdateItem_BeforeAttachmentAddEventHandler BeforeAttachmentAddEvent;

		/// <summary>
		/// SupportByVersion Outlook 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff863936.aspx </remarks>
		[SupportByVersion("Outlook", 12,14,15,16)]
		event TaskRequestUpdateItem_BeforeAttachmentPreviewEventHandler BeforeAttachmentPreviewEvent;

		/// <summary>
		/// SupportByVersion Outlook 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff866702.aspx </remarks>
		[SupportByVersion("Outlook", 12,14,15,16)]
		event TaskRequestUpdateItem_BeforeAttachmentReadEventHandler BeforeAttachmentReadEvent;

		/// <summary>
		/// SupportByVersion Outlook 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff862984.aspx </remarks>
		[SupportByVersion("Outlook", 12,14,15,16)]
		event TaskRequestUpdateItem_BeforeAttachmentWriteToTempFileEventHandler BeforeAttachmentWriteToTempFileEvent;

		/// <summary>
		/// SupportByVersion Outlook 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff868104.aspx </remarks>
		[SupportByVersion("Outlook", 12,14,15,16)]
		event TaskRequestUpdateItem_UnloadEventHandler UnloadEvent;

		/// <summary>
		/// SupportByVersion Outlook 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff868435.aspx </remarks>
		[SupportByVersion("Outlook", 12,14,15,16)]
		event TaskRequestUpdateItem_BeforeAutoSaveEventHandler BeforeAutoSaveEvent;

		/// <summary>
		/// SupportByVersion Outlook 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff860291.aspx </remarks>
		[SupportByVersion("Outlook", 14,15,16)]
		event TaskRequestUpdateItem_BeforeReadEventHandler BeforeReadEvent;

		/// <summary>
		/// SupportByVersion Outlook 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff868903.aspx </remarks>
		[SupportByVersion("Outlook", 14,15,16)]
		event TaskRequestUpdateItem_AfterWriteEventHandler AfterWriteEvent;

		/// <summary>
		/// SupportByVersion Outlook 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/jj228681.aspx </remarks>
		[SupportByVersion("Outlook", 15, 16)]
		event TaskRequestUpdateItem_ReadCompleteEventHandler ReadCompleteEvent;

        #endregion
    }
}
