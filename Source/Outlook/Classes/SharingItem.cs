using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.OutlookApi
{
	#region Delegates

	#pragma warning disable
	public delegate void SharingItem_OpenEventHandler(ref bool cancel);
	public delegate void SharingItem_CustomActionEventHandler(ICOMObject action, ICOMObject response, ref bool cancel);
	public delegate void SharingItem_CustomPropertyChangeEventHandler(string Name);
	public delegate void SharingItem_ForwardEventHandler(ICOMObject forward, ref bool cancel);
	public delegate void SharingItem_CloseEventHandler(ref bool cancel);
	public delegate void SharingItem_PropertyChangeEventHandler(string Name);
	public delegate void SharingItem_ReadEventHandler();
	public delegate void SharingItem_ReplyEventHandler(ICOMObject response, ref bool cancel);
	public delegate void SharingItem_ReplyAllEventHandler(ICOMObject response, ref bool cancel);
	public delegate void SharingItem_SendEventHandler(ref bool cancel);
	public delegate void SharingItem_WriteEventHandler(ref bool cancel);
	public delegate void SharingItem_BeforeCheckNamesEventHandler(ref bool cancel);
	public delegate void SharingItem_AttachmentAddEventHandler(NetOffice.OutlookApi.Attachment attachment);
	public delegate void SharingItem_AttachmentReadEventHandler(NetOffice.OutlookApi.Attachment attachment);
	public delegate void SharingItem_BeforeAttachmentSaveEventHandler(NetOffice.OutlookApi.Attachment attachment, ref bool cancel);
	public delegate void SharingItem_BeforeDeleteEventHandler(ICOMObject item, ref bool cancel);
	public delegate void SharingItem_AttachmentRemoveEventHandler(NetOffice.OutlookApi.Attachment attachment);
	public delegate void SharingItem_BeforeAttachmentAddEventHandler(NetOffice.OutlookApi.Attachment attachment, ref bool cancel);
	public delegate void SharingItem_BeforeAttachmentPreviewEventHandler(NetOffice.OutlookApi.Attachment attachment, ref bool cancel);
	public delegate void SharingItem_BeforeAttachmentReadEventHandler(NetOffice.OutlookApi.Attachment attachment, ref bool cancel);
	public delegate void SharingItem_BeforeAttachmentWriteToTempFileEventHandler(NetOffice.OutlookApi.Attachment attachment, ref bool cancel);
	public delegate void SharingItem_UnloadEventHandler();
	public delegate void SharingItem_BeforeAutoSaveEventHandler(ref bool cancel);
	public delegate void SharingItem_BeforeReadEventHandler();
	public delegate void SharingItem_AfterWriteEventHandler();
	public delegate void SharingItem_ReadCompleteEventHandler(ref bool cancel);
	#pragma warning restore

	#endregion

	/// <summary>
	/// CoClass SharingItem 
	/// SupportByVersion Outlook, 12,14,15,16
	/// </summary>
	/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff865852.aspx </remarks>
	[SupportByVersion("Outlook", 12,14,15,16)]
	[EntityType(EntityType.IsCoClass)]
    [ComEventContract(typeof(EventContracts.ItemEvents_10))]
	[TypeId("00061067-0000-0000-C000-000000000046")]
    public interface SharingItem : _SharingItem, IEventBinding
	{
        #region Events

        /// <summary>
        /// SupportByVersion Outlook 10 11 12 14 15,16
        /// </summary>
        ///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff868717.aspx </remarks>
        [SupportByVersion("Outlook", 10,11,12,14,15,16)]
		event SharingItem_OpenEventHandler OpenEvent;

		/// <summary>
		/// SupportByVersion Outlook 10 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff866203.aspx </remarks>
		[SupportByVersion("Outlook", 10,11,12,14,15,16)]
		event SharingItem_CustomActionEventHandler CustomActionEvent;

		/// <summary>
		/// SupportByVersion Outlook 10 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff870103.aspx </remarks>
		[SupportByVersion("Outlook", 10,11,12,14,15,16)]
		event SharingItem_CustomPropertyChangeEventHandler CustomPropertyChangeEvent;

		/// <summary>
		/// SupportByVersion Outlook 10 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff868764.aspx </remarks>
		[SupportByVersion("Outlook", 10,11,12,14,15,16)]
		event SharingItem_ForwardEventHandler ForwardEvent;

		/// <summary>
		/// SupportByVersion Outlook 10 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff860972.aspx </remarks>
		[SupportByVersion("Outlook", 10,11,12,14,15,16)]
		event SharingItem_CloseEventHandler CloseEvent;

		/// <summary>
		/// SupportByVersion Outlook 10 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff866954.aspx </remarks>
		[SupportByVersion("Outlook", 10,11,12,14,15,16)]
		event SharingItem_PropertyChangeEventHandler PropertyChangeEvent;

		/// <summary>
		/// SupportByVersion Outlook 10 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff862783.aspx </remarks>
		[SupportByVersion("Outlook", 10,11,12,14,15,16)]
		event SharingItem_ReadEventHandler ReadEvent;

		/// <summary>
		/// SupportByVersion Outlook 10 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff865602.aspx </remarks>
		[SupportByVersion("Outlook", 10,11,12,14,15,16)]
		event SharingItem_ReplyEventHandler ReplyEvent;

		/// <summary>
		/// SupportByVersion Outlook 10 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff861545.aspx </remarks>
		[SupportByVersion("Outlook", 10,11,12,14,15,16)]
		event SharingItem_ReplyAllEventHandler ReplyAllEvent;

		/// <summary>
		/// SupportByVersion Outlook 10 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff861581.aspx </remarks>
		[SupportByVersion("Outlook", 10,11,12,14,15,16)]
		event SharingItem_SendEventHandler SendEvent;

		/// <summary>
		/// SupportByVersion Outlook 10 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff862385.aspx </remarks>
		[SupportByVersion("Outlook", 10,11,12,14,15,16)]
		event SharingItem_WriteEventHandler WriteEvent;

		/// <summary>
		/// SupportByVersion Outlook 10 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff870019.aspx </remarks>
		[SupportByVersion("Outlook", 10,11,12,14,15,16)]
		event SharingItem_BeforeCheckNamesEventHandler BeforeCheckNamesEvent;

		/// <summary>
		/// SupportByVersion Outlook 10 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff868778.aspx </remarks>
		[SupportByVersion("Outlook", 10,11,12,14,15,16)]
		event SharingItem_AttachmentAddEventHandler AttachmentAddEvent;

		/// <summary>
		/// SupportByVersion Outlook 10 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff867235.aspx </remarks>
		[SupportByVersion("Outlook", 10,11,12,14,15,16)]
		event SharingItem_AttachmentReadEventHandler AttachmentReadEvent;

		/// <summary>
		/// SupportByVersion Outlook 10 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff869748.aspx </remarks>
		[SupportByVersion("Outlook", 10,11,12,14,15,16)]
		event SharingItem_BeforeAttachmentSaveEventHandler BeforeAttachmentSaveEvent;

		/// <summary>
		/// SupportByVersion Outlook 10 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff865674.aspx </remarks>
		[SupportByVersion("Outlook", 10,11,12,14,15,16)]
		event SharingItem_BeforeDeleteEventHandler BeforeDeleteEvent;

		/// <summary>
		/// SupportByVersion Outlook 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff869590.aspx </remarks>
		[SupportByVersion("Outlook", 12,14,15,16)]
		event SharingItem_AttachmentRemoveEventHandler AttachmentRemoveEvent;

		/// <summary>
		/// SupportByVersion Outlook 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff867273.aspx </remarks>
		[SupportByVersion("Outlook", 12,14,15,16)]
		event SharingItem_BeforeAttachmentAddEventHandler BeforeAttachmentAddEvent;

		/// <summary>
		/// SupportByVersion Outlook 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff869627.aspx </remarks>
		[SupportByVersion("Outlook", 12,14,15,16)]
		event SharingItem_BeforeAttachmentPreviewEventHandler BeforeAttachmentPreviewEvent;

		/// <summary>
		/// SupportByVersion Outlook 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff868931.aspx </remarks>
		[SupportByVersion("Outlook", 12,14,15,16)]
		event SharingItem_BeforeAttachmentReadEventHandler BeforeAttachmentReadEvent;

		/// <summary>
		/// SupportByVersion Outlook 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff867309.aspx </remarks>
		[SupportByVersion("Outlook", 12,14,15,16)]
		event SharingItem_BeforeAttachmentWriteToTempFileEventHandler BeforeAttachmentWriteToTempFileEvent;

		/// <summary>
		/// SupportByVersion Outlook 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff868719.aspx </remarks>
		[SupportByVersion("Outlook", 12,14,15,16)]
		event SharingItem_UnloadEventHandler UnloadEvent;

		/// <summary>
		/// SupportByVersion Outlook 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff863605.aspx </remarks>
		[SupportByVersion("Outlook", 12,14,15,16)]
		event SharingItem_BeforeAutoSaveEventHandler BeforeAutoSaveEvent;

		/// <summary>
		/// SupportByVersion Outlook 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff863722.aspx </remarks>
		[SupportByVersion("Outlook", 14,15,16)]
		event SharingItem_BeforeReadEventHandler BeforeReadEvent;

		/// <summary>
		/// SupportByVersion Outlook 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff868438.aspx </remarks>
		[SupportByVersion("Outlook", 14,15,16)]
		event SharingItem_AfterWriteEventHandler AfterWriteEvent;

		/// <summary>
		/// SupportByVersion Outlook 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/jj228066.aspx </remarks>
		[SupportByVersion("Outlook", 15, 16)]
		event SharingItem_ReadCompleteEventHandler ReadCompleteEvent;

        #endregion
    }
}
