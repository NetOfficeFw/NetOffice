using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.OutlookApi
{
	#region Delegates

	#pragma warning disable
	public delegate void JournalItem_OpenEventHandler(ref bool cancel);
	public delegate void JournalItem_CustomActionEventHandler(ICOMObject action, ICOMObject response, ref bool cancel);
	public delegate void JournalItem_CustomPropertyChangeEventHandler(string name);
	public delegate void JournalItem_ForwardEventHandler(ICOMObject forward, ref bool cancel);
	public delegate void JournalItem_CloseEventHandler(ref bool cancel);
	public delegate void JournalItem_PropertyChangeEventHandler(string name);
	public delegate void JournalItem_ReadEventHandler();
	public delegate void JournalItem_ReplyEventHandler(ICOMObject response, ref bool cancel);
	public delegate void JournalItem_ReplyAllEventHandler(ICOMObject response, ref bool cancel);
	public delegate void JournalItem_SendEventHandler(ref bool cancel);
	public delegate void JournalItem_WriteEventHandler(ref bool cancel);
	public delegate void JournalItem_BeforeCheckNamesEventHandler(ref bool cancel);
	public delegate void JournalItem_AttachmentAddEventHandler(NetOffice.OutlookApi.Attachment attachment);
	public delegate void JournalItem_AttachmentReadEventHandler(NetOffice.OutlookApi.Attachment attachment);
	public delegate void JournalItem_BeforeAttachmentSaveEventHandler(NetOffice.OutlookApi.Attachment attachment, ref bool cancel);
	public delegate void JournalItem_BeforeDeleteEventHandler(ICOMObject item, ref bool cancel);
	public delegate void JournalItem_AttachmentRemoveEventHandler(NetOffice.OutlookApi.Attachment attachment);
	public delegate void JournalItem_BeforeAttachmentAddEventHandler(NetOffice.OutlookApi.Attachment attachment, ref bool cancel);
	public delegate void JournalItem_BeforeAttachmentPreviewEventHandler(NetOffice.OutlookApi.Attachment attachment, ref bool cancel);
	public delegate void JournalItem_BeforeAttachmentReadEventHandler(NetOffice.OutlookApi.Attachment attachment, ref bool cancel);
	public delegate void JournalItem_BeforeAttachmentWriteToTempFileEventHandler(NetOffice.OutlookApi.Attachment attachment, ref bool cancel);
	public delegate void JournalItem_UnloadEventHandler();
	public delegate void JournalItem_BeforeAutoSaveEventHandler(ref bool cancel);
	public delegate void JournalItem_BeforeReadEventHandler();
	public delegate void JournalItem_AfterWriteEventHandler();
	public delegate void JournalItem_ReadCompleteEventHandler(ref bool cancel);
	#pragma warning restore

	#endregion

	/// <summary>
	/// CoClass JournalItem 
	/// SupportByVersion Outlook, 9,10,11,12,14,15,16
	/// </summary>
	/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff866277.aspx </remarks>
	[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
	[EntityType(EntityType.IsCoClass)]
    [ComEventContract(typeof(EventContracts.ItemEvents), typeof(EventContracts.ItemEvents_10))]
	[TypeId("00061037-0000-0000-C000-000000000046")]
    public interface JournalItem : _JournalItem, IEventBinding
	{
        #region Events

        /// <summary>
        /// SupportByVersion Outlook 9 10 11 12 14 15,16
        /// </summary>
        ///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff869319.aspx </remarks>
        [SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		event JournalItem_OpenEventHandler OpenEvent;

		/// <summary>
		/// SupportByVersion Outlook 9 10 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff864388.aspx </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		event JournalItem_CustomActionEventHandler CustomActionEvent;

		/// <summary>
		/// SupportByVersion Outlook 9 10 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff868830.aspx </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		event JournalItem_CustomPropertyChangeEventHandler CustomPropertyChangeEvent;

		/// <summary>
		/// SupportByVersion Outlook 9 10 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff861008.aspx </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		event JournalItem_ForwardEventHandler ForwardEvent;

		/// <summary>
		/// SupportByVersion Outlook 9 10 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff866899.aspx </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		event JournalItem_CloseEventHandler CloseEvent;

		/// <summary>
		/// SupportByVersion Outlook 9 10 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff868239.aspx </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		event JournalItem_PropertyChangeEventHandler PropertyChangeEvent;

		/// <summary>
		/// SupportByVersion Outlook 9 10 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff863372.aspx </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		event JournalItem_ReadEventHandler ReadEvent;

		/// <summary>
		/// SupportByVersion Outlook 9 10 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff861597.aspx </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		event JournalItem_ReplyEventHandler ReplyEvent;

		/// <summary>
		/// SupportByVersion Outlook 9 10 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff867342.aspx </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		event JournalItem_ReplyAllEventHandler ReplyAllEvent;

		/// <summary>
		/// SupportByVersion Outlook 9 10 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff860991.aspx </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		event JournalItem_SendEventHandler SendEvent;

		/// <summary>
		/// SupportByVersion Outlook 9 10 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff865837.aspx </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		event JournalItem_WriteEventHandler WriteEvent;

		/// <summary>
		/// SupportByVersion Outlook 9 10 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff868613.aspx </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		event JournalItem_BeforeCheckNamesEventHandler BeforeCheckNamesEvent;

		/// <summary>
		/// SupportByVersion Outlook 9 10 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff867177.aspx </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		event JournalItem_AttachmentAddEventHandler AttachmentAddEvent;

		/// <summary>
		/// SupportByVersion Outlook 9 10 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff869815.aspx </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		event JournalItem_AttachmentReadEventHandler AttachmentReadEvent;

		/// <summary>
		/// SupportByVersion Outlook 9 10 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff869199.aspx </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		event JournalItem_BeforeAttachmentSaveEventHandler BeforeAttachmentSaveEvent;

		/// <summary>
		/// SupportByVersion Outlook 10 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff863097.aspx </remarks>
		[SupportByVersion("Outlook", 10,11,12,14,15,16)]
		event JournalItem_BeforeDeleteEventHandler BeforeDeleteEvent;

		/// <summary>
		/// SupportByVersion Outlook 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff866241.aspx </remarks>
		[SupportByVersion("Outlook", 12,14,15,16)]
		event JournalItem_AttachmentRemoveEventHandler AttachmentRemoveEvent;

		/// <summary>
		/// SupportByVersion Outlook 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff868966.aspx </remarks>
		[SupportByVersion("Outlook", 12,14,15,16)]
		event JournalItem_BeforeAttachmentAddEventHandler BeforeAttachmentAddEvent;

		/// <summary>
		/// SupportByVersion Outlook 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff869700.aspx </remarks>
		[SupportByVersion("Outlook", 12,14,15,16)]
		event JournalItem_BeforeAttachmentPreviewEventHandler BeforeAttachmentPreviewEvent;

		/// <summary>
		/// SupportByVersion Outlook 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff868352.aspx </remarks>
		[SupportByVersion("Outlook", 12,14,15,16)]
		event JournalItem_BeforeAttachmentReadEventHandler BeforeAttachmentReadEvent;

		/// <summary>
		/// SupportByVersion Outlook 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff860451.aspx </remarks>
		[SupportByVersion("Outlook", 12,14,15,16)]
		event JournalItem_BeforeAttachmentWriteToTempFileEventHandler BeforeAttachmentWriteToTempFileEvent;

		/// <summary>
		/// SupportByVersion Outlook 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff864746.aspx </remarks>
		[SupportByVersion("Outlook", 12,14,15,16)]
		event JournalItem_UnloadEventHandler UnloadEvent;

		/// <summary>
		/// SupportByVersion Outlook 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff868660.aspx </remarks>
		[SupportByVersion("Outlook", 12,14,15,16)]
		event JournalItem_BeforeAutoSaveEventHandler BeforeAutoSaveEvent;

		/// <summary>
		/// SupportByVersion Outlook 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff866068.aspx </remarks>
		[SupportByVersion("Outlook", 14,15,16)]
		event JournalItem_BeforeReadEventHandler BeforeReadEvent;

		/// <summary>
		/// SupportByVersion Outlook 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff868767.aspx </remarks>
		[SupportByVersion("Outlook", 14,15,16)]
		event JournalItem_AfterWriteEventHandler AfterWriteEvent;

		/// <summary>
		/// SupportByVersion Outlook 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/jj229141.aspx </remarks>
		[SupportByVersion("Outlook", 15, 16)]
		event JournalItem_ReadCompleteEventHandler ReadCompleteEvent;

        #endregion
    }
}
