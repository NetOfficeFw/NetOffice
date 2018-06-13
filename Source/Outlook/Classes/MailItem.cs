using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.OutlookApi
{
	#region Delegates

	#pragma warning disable
	public delegate void MailItem_OpenEventHandler(ref bool cancel);
	public delegate void MailItem_CustomActionEventHandler(ICOMObject action, ICOMObject response, ref bool cancel);
	public delegate void MailItem_CustomPropertyChangeEventHandler(string name);
	public delegate void MailItem_ForwardEventHandler(ICOMObject forward, ref bool cancel);
	public delegate void MailItem_CloseEventHandler(ref bool cancel);
	public delegate void MailItem_PropertyChangeEventHandler(string name);
	public delegate void MailItem_ReadEventHandler();
	public delegate void MailItem_ReplyEventHandler(ICOMObject response, ref bool cancel);
	public delegate void MailItem_ReplyAllEventHandler(ICOMObject response, ref bool cancel);
	public delegate void MailItem_SendEventHandler(ref bool cancel);
	public delegate void MailItem_WriteEventHandler(ref bool cancel);
	public delegate void MailItem_BeforeCheckNamesEventHandler(ref bool cancel);
	public delegate void MailItem_AttachmentAddEventHandler(NetOffice.OutlookApi.Attachment attachment);
	public delegate void MailItem_AttachmentReadEventHandler(NetOffice.OutlookApi.Attachment attachment);
	public delegate void MailItem_BeforeAttachmentSaveEventHandler(NetOffice.OutlookApi.Attachment attachment, ref bool cancel);
	public delegate void MailItem_BeforeDeleteEventHandler(ICOMObject item, ref bool cancel);
	public delegate void MailItem_AttachmentRemoveEventHandler(NetOffice.OutlookApi.Attachment attachment);
	public delegate void MailItem_BeforeAttachmentAddEventHandler(NetOffice.OutlookApi.Attachment attachment, ref bool cancel);
	public delegate void MailItem_BeforeAttachmentPreviewEventHandler(NetOffice.OutlookApi.Attachment attachment, ref bool cancel);
	public delegate void MailItem_BeforeAttachmentReadEventHandler(NetOffice.OutlookApi.Attachment attachment, ref bool cancel);
	public delegate void MailItem_BeforeAttachmentWriteToTempFileEventHandler(NetOffice.OutlookApi.Attachment attachment, ref bool cancel);
	public delegate void MailItem_UnloadEventHandler();
	public delegate void MailItem_BeforeAutoSaveEventHandler(ref bool Cancel);
	public delegate void MailItem_BeforeReadEventHandler();
	public delegate void MailItem_AfterWriteEventHandler();
	public delegate void MailItem_ReadCompleteEventHandler(ref bool Cancel);
	#pragma warning restore

	#endregion

	/// <summary>
	/// CoClass MailItem 
	/// SupportByVersion Outlook, 9,10,11,12,14,15,16
	/// </summary>
	/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff861332.aspx </remarks>
	[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
	[EntityType(EntityType.IsCoClass)]
    [ComEventContract(typeof(EventContracts.ItemEvents), typeof(EventContracts.ItemEvents_10))]
	[TypeId("00061033-0000-0000-C000-000000000046")]
    public interface MailItem : _MailItem, IEventBinding
	{
        #region Events

        /// <summary>
        /// SupportByVersion Outlook 9 10 11 12 14 15,16
        /// </summary>
        ///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff865989.aspx </remarks>
        [SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		event MailItem_OpenEventHandler OpenEvent;

		/// <summary>
		/// SupportByVersion Outlook 9 10 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff862186.aspx </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		event MailItem_CustomActionEventHandler CustomActionEvent;

		/// <summary>
		/// SupportByVersion Outlook 9 10 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff865310.aspx </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		event MailItem_CustomPropertyChangeEventHandler CustomPropertyChangeEvent;

		/// <summary>
		/// SupportByVersion Outlook 9 10 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff862702.aspx </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		event MailItem_ForwardEventHandler ForwardEvent;

		/// <summary>
		/// SupportByVersion Outlook 9 10 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff867865.aspx </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		event MailItem_CloseEventHandler CloseEvent;

		/// <summary>
		/// SupportByVersion Outlook 9 10 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff866739.aspx </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		event MailItem_PropertyChangeEventHandler PropertyChangeEvent;

		/// <summary>
		/// SupportByVersion Outlook 9 10 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff869872.aspx </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		event MailItem_ReadEventHandler ReadEvent;

		/// <summary>
		/// SupportByVersion Outlook 9 10 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff860938.aspx </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		event MailItem_ReplyEventHandler ReplyEvent;

		/// <summary>
		/// SupportByVersion Outlook 9 10 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff869905.aspx </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		event MailItem_ReplyAllEventHandler ReplyAllEvent;

		/// <summary>
		/// SupportByVersion Outlook 9 10 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff865379.aspx </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		event MailItem_SendEventHandler SendEvent;

		/// <summary>
		/// SupportByVersion Outlook 9 10 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff868664.aspx </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		event MailItem_WriteEventHandler WriteEvent;

		/// <summary>
		/// SupportByVersion Outlook 9 10 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff870099.aspx </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		event MailItem_BeforeCheckNamesEventHandler BeforeCheckNamesEvent;

		/// <summary>
		/// SupportByVersion Outlook 9 10 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff868540.aspx </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		event MailItem_AttachmentAddEventHandler AttachmentAddEvent;

		/// <summary>
		/// SupportByVersion Outlook 9 10 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff868187.aspx </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		event MailItem_AttachmentReadEventHandler AttachmentReadEvent;

		/// <summary>
		/// SupportByVersion Outlook 9 10 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff868639.aspx </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		event MailItem_BeforeAttachmentSaveEventHandler BeforeAttachmentSaveEvent;

		/// <summary>
		/// SupportByVersion Outlook 10 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff861266.aspx </remarks>
		[SupportByVersion("Outlook", 10,11,12,14,15,16)]
		event MailItem_BeforeDeleteEventHandler BeforeDeleteEvent;

		/// <summary>
		/// SupportByVersion Outlook 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff863737.aspx </remarks>
		[SupportByVersion("Outlook", 12,14,15,16)]
		event MailItem_AttachmentRemoveEventHandler AttachmentRemoveEvent;

		/// <summary>
		/// SupportByVersion Outlook 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff869219.aspx </remarks>
		[SupportByVersion("Outlook", 12,14,15,16)]
		event MailItem_BeforeAttachmentAddEventHandler BeforeAttachmentAddEvent;

		/// <summary>
		/// SupportByVersion Outlook 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff862669.aspx </remarks>
		[SupportByVersion("Outlook", 12,14,15,16)]
		event MailItem_BeforeAttachmentPreviewEventHandler BeforeAttachmentPreviewEvent;

		/// <summary>
		/// SupportByVersion Outlook 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff860316.aspx </remarks>
		[SupportByVersion("Outlook", 12,14,15,16)]
		event MailItem_BeforeAttachmentReadEventHandler BeforeAttachmentReadEvent;

		/// <summary>
		/// SupportByVersion Outlook 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff870101.aspx </remarks>
		[SupportByVersion("Outlook", 12,14,15,16)]
		event MailItem_BeforeAttachmentWriteToTempFileEventHandler BeforeAttachmentWriteToTempFileEvent;

		/// <summary>
		/// SupportByVersion Outlook 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff868564.aspx </remarks>
		[SupportByVersion("Outlook", 12,14,15,16)]
		event MailItem_UnloadEventHandler UnloadEvent;

		/// <summary>
		/// SupportByVersion Outlook 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff860949.aspx </remarks>
		[SupportByVersion("Outlook", 12,14,15,16)]
		event MailItem_BeforeAutoSaveEventHandler BeforeAutoSaveEvent;

		/// <summary>
		/// SupportByVersion Outlook 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff869497.aspx </remarks>
		[SupportByVersion("Outlook", 14,15,16)]
		event MailItem_BeforeReadEventHandler BeforeReadEvent;

		/// <summary>
		/// SupportByVersion Outlook 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff869691.aspx </remarks>
		[SupportByVersion("Outlook", 14,15,16)]
		event MailItem_AfterWriteEventHandler AfterWriteEvent;

		/// <summary>
		/// SupportByVersion Outlook 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/jj228313.aspx </remarks>
		[SupportByVersion("Outlook", 15, 16)]
		event MailItem_ReadCompleteEventHandler ReadCompleteEvent;

        #endregion
    }
}
