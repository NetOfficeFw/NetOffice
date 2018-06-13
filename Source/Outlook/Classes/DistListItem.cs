using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.OutlookApi
{
	#region Delegates

	#pragma warning disable
	public delegate void DistListItem_OpenEventHandler(ref bool cancel);
	public delegate void DistListItem_CustomActionEventHandler(ICOMObject action, ICOMObject response, ref bool cancel);
	public delegate void DistListItem_CustomPropertyChangeEventHandler(string name);
	public delegate void DistListItem_ForwardEventHandler(ICOMObject forward, ref bool cancel);
	public delegate void DistListItem_CloseEventHandler(ref bool cancel);
	public delegate void DistListItem_PropertyChangeEventHandler(string Name);
	public delegate void DistListItem_ReadEventHandler();
	public delegate void DistListItem_ReplyEventHandler(ICOMObject response, ref bool cancel);
	public delegate void DistListItem_ReplyAllEventHandler(ICOMObject response, ref bool cancel);
	public delegate void DistListItem_SendEventHandler(ref bool cancel);
	public delegate void DistListItem_WriteEventHandler(ref bool cancel);
	public delegate void DistListItem_BeforeCheckNamesEventHandler(ref bool cancel);
	public delegate void DistListItem_AttachmentAddEventHandler(NetOffice.OutlookApi.Attachment attachment);
	public delegate void DistListItem_AttachmentReadEventHandler(NetOffice.OutlookApi.Attachment attachment);
	public delegate void DistListItem_BeforeAttachmentSaveEventHandler(NetOffice.OutlookApi.Attachment attachment, ref bool cancel);
	public delegate void DistListItem_BeforeDeleteEventHandler(ICOMObject item, ref bool cancel);
	public delegate void DistListItem_AttachmentRemoveEventHandler(NetOffice.OutlookApi.Attachment attachment);
	public delegate void DistListItem_BeforeAttachmentAddEventHandler(NetOffice.OutlookApi.Attachment attachment, ref bool cancel);
	public delegate void DistListItem_BeforeAttachmentPreviewEventHandler(NetOffice.OutlookApi.Attachment attachment, ref bool cancel);
	public delegate void DistListItem_BeforeAttachmentReadEventHandler(NetOffice.OutlookApi.Attachment attachment, ref bool cancel);
	public delegate void DistListItem_BeforeAttachmentWriteToTempFileEventHandler(NetOffice.OutlookApi.Attachment attachment, ref bool cancel);
	public delegate void DistListItem_UnloadEventHandler();
	public delegate void DistListItem_BeforeAutoSaveEventHandler(ref bool cancel);
	public delegate void DistListItem_BeforeReadEventHandler();
	public delegate void DistListItem_AfterWriteEventHandler();
	public delegate void DistListItem_ReadCompleteEventHandler(ref bool cancel);
	#pragma warning restore

	#endregion

	/// <summary>
	/// CoClass DistListItem 
	/// SupportByVersion Outlook, 9,10,11,12,14,15,16
	/// </summary>
	/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff860361.aspx </remarks>
	[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
	[EntityType(EntityType.IsCoClass)]
    [ComEventContract(typeof(EventContracts.ItemEvents), typeof(EventContracts.ItemEvents_10))]
	[TypeId("0006103C-0000-0000-C000-000000000046")]
    public interface DistListItem : _DistListItem, IEventBinding
	{
        #region Events

        /// <summary>
        /// SupportByVersion Outlook 9 10 11 12 14 15,16
        /// </summary>
        ///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff865402.aspx </remarks>
        [SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		event DistListItem_OpenEventHandler OpenEvent;

		/// <summary>
		/// SupportByVersion Outlook 9 10 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff869156.aspx </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		event DistListItem_CustomActionEventHandler CustomActionEvent;

		/// <summary>
		/// SupportByVersion Outlook 9 10 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff867633.aspx </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		event DistListItem_CustomPropertyChangeEventHandler CustomPropertyChangeEvent;

		/// <summary>
		/// SupportByVersion Outlook 9 10 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff862716.aspx </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		event DistListItem_ForwardEventHandler ForwardEvent;

		/// <summary>
		/// SupportByVersion Outlook 9 10 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff868458.aspx </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		event DistListItem_CloseEventHandler CloseEvent;

		/// <summary>
		/// SupportByVersion Outlook 9 10 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff867815.aspx </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		event DistListItem_PropertyChangeEventHandler PropertyChangeEvent;

		/// <summary>
		/// SupportByVersion Outlook 9 10 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff865312.aspx </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		event DistListItem_ReadEventHandler ReadEvent;

		/// <summary>
		/// SupportByVersion Outlook 9 10 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff867332.aspx </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		event DistListItem_ReplyEventHandler ReplyEvent;

		/// <summary>
		/// SupportByVersion Outlook 9 10 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff865842.aspx </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		event DistListItem_ReplyAllEventHandler ReplyAllEvent;

		/// <summary>
		/// SupportByVersion Outlook 9 10 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff867634.aspx </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		event DistListItem_SendEventHandler SendEvent;

		/// <summary>
		/// SupportByVersion Outlook 9 10 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff869093.aspx </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		event DistListItem_WriteEventHandler WriteEvent;

		/// <summary>
		/// SupportByVersion Outlook 9 10 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff864770.aspx </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		event DistListItem_BeforeCheckNamesEventHandler BeforeCheckNamesEvent;

		/// <summary>
		/// SupportByVersion Outlook 9 10 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff867823.aspx </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		event DistListItem_AttachmentAddEventHandler AttachmentAddEvent;

		/// <summary>
		/// SupportByVersion Outlook 9 10 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff861918.aspx </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		event DistListItem_AttachmentReadEventHandler AttachmentReadEvent;

		/// <summary>
		/// SupportByVersion Outlook 9 10 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff865596.aspx </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		event DistListItem_BeforeAttachmentSaveEventHandler BeforeAttachmentSaveEvent;

		/// <summary>
		/// SupportByVersion Outlook 10 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff860709.aspx </remarks>
		[SupportByVersion("Outlook", 10,11,12,14,15,16)]
		event DistListItem_BeforeDeleteEventHandler BeforeDeleteEvent;

		/// <summary>
		/// SupportByVersion Outlook 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff860675.aspx </remarks>
		[SupportByVersion("Outlook", 12,14,15,16)]
		event DistListItem_AttachmentRemoveEventHandler AttachmentRemoveEvent;

		/// <summary>
		/// SupportByVersion Outlook 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff867883.aspx </remarks>
		[SupportByVersion("Outlook", 12,14,15,16)]
		event DistListItem_BeforeAttachmentAddEventHandler BeforeAttachmentAddEvent;

		/// <summary>
		/// SupportByVersion Outlook 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff869950.aspx </remarks>
		[SupportByVersion("Outlook", 12,14,15,16)]
		event DistListItem_BeforeAttachmentPreviewEventHandler BeforeAttachmentPreviewEvent;

		/// <summary>
		/// SupportByVersion Outlook 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff870034.aspx </remarks>
		[SupportByVersion("Outlook", 12,14,15,16)]
		event DistListItem_BeforeAttachmentReadEventHandler BeforeAttachmentReadEvent;

		/// <summary>
		/// SupportByVersion Outlook 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff865268.aspx </remarks>
		[SupportByVersion("Outlook", 12,14,15,16)]
		event DistListItem_BeforeAttachmentWriteToTempFileEventHandler BeforeAttachmentWriteToTempFileEvent;

		/// <summary>
		/// SupportByVersion Outlook 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff862480.aspx </remarks>
		[SupportByVersion("Outlook", 12,14,15,16)]
		event DistListItem_UnloadEventHandler UnloadEvent;

		/// <summary>
		/// SupportByVersion Outlook 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff868780.aspx </remarks>
		[SupportByVersion("Outlook", 12,14,15,16)]
		event DistListItem_BeforeAutoSaveEventHandler BeforeAutoSaveEvent;

		/// <summary>
		/// SupportByVersion Outlook 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff864704.aspx </remarks>
		[SupportByVersion("Outlook", 14,15,16)]
		event DistListItem_BeforeReadEventHandler BeforeReadEvent;

		/// <summary>
		/// SupportByVersion Outlook 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff868457.aspx </remarks>
		[SupportByVersion("Outlook", 14,15,16)]
		event DistListItem_AfterWriteEventHandler AfterWriteEvent;

		/// <summary>
		/// SupportByVersion Outlook 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/jj227209.aspx </remarks>
		[SupportByVersion("Outlook", 15, 16)]
		event DistListItem_ReadCompleteEventHandler ReadCompleteEvent;

        #endregion
    }
}
