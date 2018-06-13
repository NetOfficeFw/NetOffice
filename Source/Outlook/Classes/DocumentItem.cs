using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.OutlookApi
{
	#region Delegates

	#pragma warning disable
	public delegate void DocumentItem_OpenEventHandler(ref bool cancel);
	public delegate void DocumentItem_CustomActionEventHandler(ICOMObject action, ICOMObject response, ref bool cancel);
	public delegate void DocumentItem_CustomPropertyChangeEventHandler(string name);
	public delegate void DocumentItem_ForwardEventHandler(ICOMObject forward, ref bool cancel);
	public delegate void DocumentItem_CloseEventHandler(ref bool cancel);
	public delegate void DocumentItem_PropertyChangeEventHandler(string name);
	public delegate void DocumentItem_ReadEventHandler();
	public delegate void DocumentItem_ReplyEventHandler(ICOMObject response, ref bool cancel);
	public delegate void DocumentItem_ReplyAllEventHandler(ICOMObject Response, ref bool cancel);
	public delegate void DocumentItem_SendEventHandler(ref bool cancel);
	public delegate void DocumentItem_WriteEventHandler(ref bool cancel);
	public delegate void DocumentItem_BeforeCheckNamesEventHandler(ref bool cancel);
	public delegate void DocumentItem_AttachmentAddEventHandler(NetOffice.OutlookApi.Attachment attachment);
	public delegate void DocumentItem_AttachmentReadEventHandler(NetOffice.OutlookApi.Attachment attachment);
	public delegate void DocumentItem_BeforeAttachmentSaveEventHandler(NetOffice.OutlookApi.Attachment attachment, ref bool cancel);
	public delegate void DocumentItem_BeforeDeleteEventHandler(ICOMObject item, ref bool cancel);
	public delegate void DocumentItem_AttachmentRemoveEventHandler(NetOffice.OutlookApi.Attachment attachment);
	public delegate void DocumentItem_BeforeAttachmentAddEventHandler(NetOffice.OutlookApi.Attachment attachment, ref bool cancel);
	public delegate void DocumentItem_BeforeAttachmentPreviewEventHandler(NetOffice.OutlookApi.Attachment attachment, ref bool cancel);
	public delegate void DocumentItem_BeforeAttachmentReadEventHandler(NetOffice.OutlookApi.Attachment attachment, ref bool cancel);
	public delegate void DocumentItem_BeforeAttachmentWriteToTempFileEventHandler(NetOffice.OutlookApi.Attachment attachment, ref bool cancel);
	public delegate void DocumentItem_UnloadEventHandler();
	public delegate void DocumentItem_BeforeAutoSaveEventHandler(ref bool cancel);
	public delegate void DocumentItem_BeforeReadEventHandler();
	public delegate void DocumentItem_AfterWriteEventHandler();
	public delegate void DocumentItem_ReadCompleteEventHandler(ref bool cancel);
	#pragma warning restore

	#endregion

	/// <summary>
	/// CoClass DocumentItem 
	/// SupportByVersion Outlook, 9,10,11,12,14,15,16
	/// </summary>
	/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff866928.aspx </remarks>
	[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
	[EntityType(EntityType.IsCoClass)]
    [ComEventContract(typeof(EventContracts.ItemEvents), typeof(EventContracts.ItemEvents_10))]
	[TypeId("00061061-0000-0000-C000-000000000046")]
    public interface DocumentItem : _DocumentItem, IEventBinding
	{
        #region Events

        /// <summary>
        /// SupportByVersion Outlook 9 10 11 12 14 15,16
        /// </summary>
        ///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff869663.aspx </remarks>
        [SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		event DocumentItem_OpenEventHandler OpenEvent;

		/// <summary>
		/// SupportByVersion Outlook 9 10 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff869809.aspx </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		event DocumentItem_CustomActionEventHandler CustomActionEvent;

		/// <summary>
		/// SupportByVersion Outlook 9 10 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff861289.aspx </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		event DocumentItem_CustomPropertyChangeEventHandler CustomPropertyChangeEvent;

		/// <summary>
		/// SupportByVersion Outlook 9 10 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff863623.aspx </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		event DocumentItem_ForwardEventHandler ForwardEvent;

		/// <summary>
		/// SupportByVersion Outlook 9 10 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff861319.aspx </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		event DocumentItem_CloseEventHandler CloseEvent;

		/// <summary>
		/// SupportByVersion Outlook 9 10 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff869749.aspx </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		event DocumentItem_PropertyChangeEventHandler PropertyChangeEvent;

		/// <summary>
		/// SupportByVersion Outlook 9 10 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff869425.aspx </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		event DocumentItem_ReadEventHandler ReadEvent;

		/// <summary>
		/// SupportByVersion Outlook 9 10 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff862733.aspx </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		event DocumentItem_ReplyEventHandler ReplyEvent;

		/// <summary>
		/// SupportByVersion Outlook 9 10 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff868691.aspx </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		event DocumentItem_ReplyAllEventHandler ReplyAllEvent;

		/// <summary>
		/// SupportByVersion Outlook 9 10 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff867100.aspx </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		event DocumentItem_SendEventHandler SendEvent;

		/// <summary>
		/// SupportByVersion Outlook 9 10 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff868539.aspx </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		event DocumentItem_WriteEventHandler WriteEvent;

		/// <summary>
		/// SupportByVersion Outlook 9 10 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff860659.aspx </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		event DocumentItem_BeforeCheckNamesEventHandler BeforeCheckNamesEvent;

		/// <summary>
		/// SupportByVersion Outlook 9 10 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff862370.aspx </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		event DocumentItem_AttachmentAddEventHandler AttachmentAddEvent;

		/// <summary>
		/// SupportByVersion Outlook 9 10 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff864394.aspx </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		event DocumentItem_AttachmentReadEventHandler AttachmentReadEvent;

		/// <summary>
		/// SupportByVersion Outlook 9 10 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff865084.aspx </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		event DocumentItem_BeforeAttachmentSaveEventHandler BeforeAttachmentSaveEvent;

		/// <summary>
		/// SupportByVersion Outlook 10 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff866473.aspx </remarks>
		[SupportByVersion("Outlook", 10,11,12,14,15,16)]
		event DocumentItem_BeforeDeleteEventHandler BeforeDeleteEvent;

		/// <summary>
		/// SupportByVersion Outlook 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff869067.aspx </remarks>
		[SupportByVersion("Outlook", 12,14,15,16)]
		event DocumentItem_AttachmentRemoveEventHandler AttachmentRemoveEvent;

		/// <summary>
		/// SupportByVersion Outlook 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff869141.aspx </remarks>
		[SupportByVersion("Outlook", 12,14,15,16)]
		event DocumentItem_BeforeAttachmentAddEventHandler BeforeAttachmentAddEvent;

		/// <summary>
		/// SupportByVersion Outlook 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff866045.aspx </remarks>
		[SupportByVersion("Outlook", 12,14,15,16)]
		event DocumentItem_BeforeAttachmentPreviewEventHandler BeforeAttachmentPreviewEvent;

		/// <summary>
		/// SupportByVersion Outlook 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff862392.aspx </remarks>
		[SupportByVersion("Outlook", 12,14,15,16)]
		event DocumentItem_BeforeAttachmentReadEventHandler BeforeAttachmentReadEvent;

		/// <summary>
		/// SupportByVersion Outlook 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff860740.aspx </remarks>
		[SupportByVersion("Outlook", 12,14,15,16)]
		event DocumentItem_BeforeAttachmentWriteToTempFileEventHandler BeforeAttachmentWriteToTempFileEvent;

		/// <summary>
		/// SupportByVersion Outlook 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff869637.aspx </remarks>
		[SupportByVersion("Outlook", 12,14,15,16)]
		event DocumentItem_UnloadEventHandler UnloadEvent;

		/// <summary>
		/// SupportByVersion Outlook 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff863656.aspx </remarks>
		[SupportByVersion("Outlook", 12,14,15,16)]
		event DocumentItem_BeforeAutoSaveEventHandler BeforeAutoSaveEvent;

		/// <summary>
		/// SupportByVersion Outlook 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff865392.aspx </remarks>
		[SupportByVersion("Outlook", 14,15,16)]
		event DocumentItem_BeforeReadEventHandler BeforeReadEvent;

		/// <summary>
		/// SupportByVersion Outlook 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff870043.aspx </remarks>
		[SupportByVersion("Outlook", 14,15,16)]
		event DocumentItem_AfterWriteEventHandler AfterWriteEvent;

		/// <summary>
		/// SupportByVersion Outlook 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/jj228964.aspx </remarks>
		[SupportByVersion("Outlook", 15, 16)]
		event DocumentItem_ReadCompleteEventHandler ReadCompleteEvent;

        #endregion
    }
}
