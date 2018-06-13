using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.OutlookApi
{
	#region Delegates

	#pragma warning disable
	public delegate void ReportItem_OpenEventHandler(ref bool cancel);
	public delegate void ReportItem_CustomActionEventHandler(ICOMObject action, ICOMObject response, ref bool cancel);
	public delegate void ReportItem_CustomPropertyChangeEventHandler(string Name);
	public delegate void ReportItem_ForwardEventHandler(ICOMObject forward, ref bool cancel);
	public delegate void ReportItem_CloseEventHandler(ref bool cancel);
	public delegate void ReportItem_PropertyChangeEventHandler(string Name);
	public delegate void ReportItem_ReadEventHandler();
	public delegate void ReportItem_ReplyEventHandler(ICOMObject response, ref bool cancel);
	public delegate void ReportItem_ReplyAllEventHandler(ICOMObject response, ref bool cancel);
	public delegate void ReportItem_SendEventHandler(ref bool cancel);
	public delegate void ReportItem_WriteEventHandler(ref bool cancel);
	public delegate void ReportItem_BeforeCheckNamesEventHandler(ref bool cancel);
	public delegate void ReportItem_AttachmentAddEventHandler(NetOffice.OutlookApi.Attachment attachment);
	public delegate void ReportItem_AttachmentReadEventHandler(NetOffice.OutlookApi.Attachment attachment);
	public delegate void ReportItem_BeforeAttachmentSaveEventHandler(NetOffice.OutlookApi.Attachment attachment, ref bool cancel);
	public delegate void ReportItem_BeforeDeleteEventHandler(ICOMObject item, ref bool cancel);
	public delegate void ReportItem_AttachmentRemoveEventHandler(NetOffice.OutlookApi.Attachment attachment);
	public delegate void ReportItem_BeforeAttachmentAddEventHandler(NetOffice.OutlookApi.Attachment attachment, ref bool cancel);
	public delegate void ReportItem_BeforeAttachmentPreviewEventHandler(NetOffice.OutlookApi.Attachment attachment, ref bool cancel);
	public delegate void ReportItem_BeforeAttachmentReadEventHandler(NetOffice.OutlookApi.Attachment attachment, ref bool cancel);
	public delegate void ReportItem_BeforeAttachmentWriteToTempFileEventHandler(NetOffice.OutlookApi.Attachment attachment, ref bool cancel);
	public delegate void ReportItem_UnloadEventHandler();
	public delegate void ReportItem_BeforeAutoSaveEventHandler(ref bool cancel);
	public delegate void ReportItem_BeforeReadEventHandler();
	public delegate void ReportItem_AfterWriteEventHandler();
	public delegate void ReportItem_ReadCompleteEventHandler(ref bool cancel);
	#pragma warning restore

	#endregion

	/// <summary>
	/// CoClass ReportItem 
	/// SupportByVersion Outlook, 9,10,11,12,14,15,16
	/// </summary>
	/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff861605.aspx </remarks>
	[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
	[EntityType(EntityType.IsCoClass)]
    [ComEventContract(typeof(EventContracts.ItemEvents), typeof(EventContracts.ItemEvents_10))]
	[TypeId("00061035-0000-0000-C000-000000000046")]
    public interface ReportItem : _ReportItem, IEventBinding
	{
        #region Events

        /// <summary>
        /// SupportByVersion Outlook 9 10 11 12 14 15,16
        /// </summary>
        ///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff869937.aspx </remarks>
        [SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		event ReportItem_OpenEventHandler OpenEvent;

		/// <summary>
		/// SupportByVersion Outlook 9 10 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff863307.aspx </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		event ReportItem_CustomActionEventHandler CustomActionEvent;

		/// <summary>
		/// SupportByVersion Outlook 9 10 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff867514.aspx </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		event ReportItem_CustomPropertyChangeEventHandler CustomPropertyChangeEvent;

		/// <summary>
		/// SupportByVersion Outlook 9 10 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff865675.aspx </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		event ReportItem_ForwardEventHandler ForwardEvent;

		/// <summary>
		/// SupportByVersion Outlook 9 10 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff869256.aspx </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		event ReportItem_CloseEventHandler CloseEvent;

		/// <summary>
		/// SupportByVersion Outlook 9 10 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff865664.aspx </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		event ReportItem_PropertyChangeEventHandler PropertyChangeEvent;

		/// <summary>
		/// SupportByVersion Outlook 9 10 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff866930.aspx </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		event ReportItem_ReadEventHandler ReadEvent;

		/// <summary>
		/// SupportByVersion Outlook 9 10 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff869586.aspx </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		event ReportItem_ReplyEventHandler ReplyEvent;

		/// <summary>
		/// SupportByVersion Outlook 9 10 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff868683.aspx </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		event ReportItem_ReplyAllEventHandler ReplyAllEvent;

		/// <summary>
		/// SupportByVersion Outlook 9 10 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff868459.aspx </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		event ReportItem_SendEventHandler SendEvent;

		/// <summary>
		/// SupportByVersion Outlook 9 10 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff861591.aspx </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		event ReportItem_WriteEventHandler WriteEvent;

		/// <summary>
		/// SupportByVersion Outlook 9 10 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff868270.aspx </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		event ReportItem_BeforeCheckNamesEventHandler BeforeCheckNamesEvent;

		/// <summary>
		/// SupportByVersion Outlook 9 10 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff869625.aspx </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		event ReportItem_AttachmentAddEventHandler AttachmentAddEvent;

		/// <summary>
		/// SupportByVersion Outlook 9 10 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff861603.aspx </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		event ReportItem_AttachmentReadEventHandler AttachmentReadEvent;

		/// <summary>
		/// SupportByVersion Outlook 9 10 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff863942.aspx </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		event ReportItem_BeforeAttachmentSaveEventHandler BeforeAttachmentSaveEvent;

		/// <summary>
		/// SupportByVersion Outlook 10 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff863049.aspx </remarks>
		[SupportByVersion("Outlook", 10,11,12,14,15,16)]
		event ReportItem_BeforeDeleteEventHandler BeforeDeleteEvent;

		/// <summary>
		/// SupportByVersion Outlook 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff868194.aspx </remarks>
		[SupportByVersion("Outlook", 12,14,15,16)]
		event ReportItem_AttachmentRemoveEventHandler AttachmentRemoveEvent;

		/// <summary>
		/// SupportByVersion Outlook 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff869057.aspx </remarks>
		[SupportByVersion("Outlook", 12,14,15,16)]
		event ReportItem_BeforeAttachmentAddEventHandler BeforeAttachmentAddEvent;

		/// <summary>
		/// SupportByVersion Outlook 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff861248.aspx </remarks>
		[SupportByVersion("Outlook", 12,14,15,16)]
		event ReportItem_BeforeAttachmentPreviewEventHandler BeforeAttachmentPreviewEvent;

		/// <summary>
		/// SupportByVersion Outlook 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff865982.aspx </remarks>
		[SupportByVersion("Outlook", 12,14,15,16)]
		event ReportItem_BeforeAttachmentReadEventHandler BeforeAttachmentReadEvent;

		/// <summary>
		/// SupportByVersion Outlook 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff868974.aspx </remarks>
		[SupportByVersion("Outlook", 12,14,15,16)]
		event ReportItem_BeforeAttachmentWriteToTempFileEventHandler BeforeAttachmentWriteToTempFileEvent;

		/// <summary>
		/// SupportByVersion Outlook 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff867820.aspx </remarks>
		[SupportByVersion("Outlook", 12,14,15,16)]
		event ReportItem_UnloadEventHandler UnloadEvent;

		/// <summary>
		/// SupportByVersion Outlook 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff868950.aspx </remarks>
		[SupportByVersion("Outlook", 12,14,15,16)]
		event ReportItem_BeforeAutoSaveEventHandler BeforeAutoSaveEvent;

		/// <summary>
		/// SupportByVersion Outlook 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff869457.aspx </remarks>
		[SupportByVersion("Outlook", 14,15,16)]
		event ReportItem_BeforeReadEventHandler BeforeReadEvent;

		/// <summary>
		/// SupportByVersion Outlook 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff868343.aspx </remarks>
		[SupportByVersion("Outlook", 14,15,16)]
		event ReportItem_AfterWriteEventHandler AfterWriteEvent;

		/// <summary>
		/// SupportByVersion Outlook 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/jj231496.aspx </remarks>
		[SupportByVersion("Outlook", 15, 16)]
		event ReportItem_ReadCompleteEventHandler ReadCompleteEvent;

        #endregion
    }
}
