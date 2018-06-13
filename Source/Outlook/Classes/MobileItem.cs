using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.OutlookApi
{
	#region Delegates

	#pragma warning disable
	public delegate void MobileItem_OpenEventHandler(ref bool cancel);
	public delegate void MobileItem_CustomActionEventHandler(ICOMObject action, ICOMObject response, ref bool cancel);
	public delegate void MobileItem_CustomPropertyChangeEventHandler(string name);
	public delegate void MobileItem_ForwardEventHandler(ICOMObject forward, ref bool cancel);
	public delegate void MobileItem_CloseEventHandler(ref bool cancel);
	public delegate void MobileItem_PropertyChangeEventHandler(string name);
	public delegate void MobileItem_ReadEventHandler();
	public delegate void MobileItem_ReplyEventHandler(ICOMObject response, ref bool cancel);
	public delegate void MobileItem_ReplyAllEventHandler(ICOMObject response, ref bool cancel);
	public delegate void MobileItem_SendEventHandler(ref bool cancel);
	public delegate void MobileItem_WriteEventHandler(ref bool cancel);
	public delegate void MobileItem_BeforeCheckNamesEventHandler(ref bool cancel);
	public delegate void MobileItem_AttachmentAddEventHandler(NetOffice.OutlookApi.Attachment attachment);
	public delegate void MobileItem_AttachmentReadEventHandler(NetOffice.OutlookApi.Attachment attachment);
	public delegate void MobileItem_BeforeAttachmentSaveEventHandler(NetOffice.OutlookApi.Attachment attachment, ref bool cancel);
	public delegate void MobileItem_BeforeDeleteEventHandler(ICOMObject item, ref bool cancel);
	public delegate void MobileItem_AttachmentRemoveEventHandler(NetOffice.OutlookApi.Attachment attachment);
	public delegate void MobileItem_BeforeAttachmentAddEventHandler(NetOffice.OutlookApi.Attachment attachment, ref bool cancel);
	public delegate void MobileItem_BeforeAttachmentPreviewEventHandler(NetOffice.OutlookApi.Attachment attachment, ref bool cancel);
	public delegate void MobileItem_BeforeAttachmentReadEventHandler(NetOffice.OutlookApi.Attachment attachment, ref bool cancel);
	public delegate void MobileItem_BeforeAttachmentWriteToTempFileEventHandler(NetOffice.OutlookApi.Attachment attachment, ref bool cancel);
	public delegate void MobileItem_UnloadEventHandler();
	public delegate void MobileItem_BeforeAutoSaveEventHandler(ref bool cancel);
	public delegate void MobileItem_BeforeReadEventHandler();
	public delegate void MobileItem_AfterWriteEventHandler();
	public delegate void MobileItem_ReadCompleteEventHandler(ref bool cancel);
	#pragma warning restore

	#endregion

	/// <summary>
	/// CoClass MobileItem 
	/// SupportByVersion Outlook, 14,15,16
	/// </summary>
	[SupportByVersion("Outlook", 14,15,16)]
	[EntityType(EntityType.IsCoClass)]
    [ComEventContract(typeof(EventContracts.ItemEvents_10))]
	[TypeId("000610FE-0000-0000-C000-000000000046")]    
    public interface MobileItem : _MobileItem, IEventBinding
	{
        #region Events

        /// <summary>
        /// SupportByVersion Outlook 10 11 12 14 15,16
        /// </summary>
        [SupportByVersion("Outlook", 10,11,12,14,15,16)]
		event MobileItem_OpenEventHandler OpenEvent;

		/// <summary>
		/// SupportByVersion Outlook 10 11 12 14 15,16
		/// </summary>
		[SupportByVersion("Outlook", 10,11,12,14,15,16)]
		event MobileItem_CustomActionEventHandler CustomActionEvent;

		/// <summary>
		/// SupportByVersion Outlook 10 11 12 14 15,16
		/// </summary>
		[SupportByVersion("Outlook", 10,11,12,14,15,16)]
		event MobileItem_CustomPropertyChangeEventHandler CustomPropertyChangeEvent;

		/// <summary>
		/// SupportByVersion Outlook 10 11 12 14 15,16
		/// </summary>
		[SupportByVersion("Outlook", 10,11,12,14,15,16)]
		event MobileItem_ForwardEventHandler ForwardEvent;

		/// <summary>
		/// SupportByVersion Outlook 10 11 12 14 15,16
		/// </summary>
		[SupportByVersion("Outlook", 10,11,12,14,15,16)]
		event MobileItem_CloseEventHandler CloseEvent;

		/// <summary>
		/// SupportByVersion Outlook 10 11 12 14 15,16
		/// </summary>
		[SupportByVersion("Outlook", 10,11,12,14,15,16)]
		event MobileItem_PropertyChangeEventHandler PropertyChangeEvent;

		/// <summary>
		/// SupportByVersion Outlook 10 11 12 14 15,16
		/// </summary>
		[SupportByVersion("Outlook", 10,11,12,14,15,16)]
		event MobileItem_ReadEventHandler ReadEvent;

		/// <summary>
		/// SupportByVersion Outlook 10 11 12 14 15,16
		/// </summary>
		[SupportByVersion("Outlook", 10,11,12,14,15,16)]
		event MobileItem_ReplyEventHandler ReplyEvent;

		/// <summary>
		/// SupportByVersion Outlook 10 11 12 14 15,16
		/// </summary>
		[SupportByVersion("Outlook", 10,11,12,14,15,16)]
		event MobileItem_ReplyAllEventHandler ReplyAllEvent;

		/// <summary>
		/// SupportByVersion Outlook 10 11 12 14 15,16
		/// </summary>
		[SupportByVersion("Outlook", 10,11,12,14,15,16)]
		event MobileItem_SendEventHandler SendEvent;

		/// <summary>
		/// SupportByVersion Outlook 10 11 12 14 15,16
		/// </summary>
		[SupportByVersion("Outlook", 10,11,12,14,15,16)]
		event MobileItem_WriteEventHandler WriteEvent;

		/// <summary>
		/// SupportByVersion Outlook 10 11 12 14 15,16
		/// </summary>
		[SupportByVersion("Outlook", 10,11,12,14,15,16)]
		event MobileItem_BeforeCheckNamesEventHandler BeforeCheckNamesEvent;

		/// <summary>
		/// SupportByVersion Outlook 10 11 12 14 15,16
		/// </summary>
		[SupportByVersion("Outlook", 10,11,12,14,15,16)]
		event MobileItem_AttachmentAddEventHandler AttachmentAddEvent;

		/// <summary>
		/// SupportByVersion Outlook 10 11 12 14 15,16
		/// </summary>
		[SupportByVersion("Outlook", 10,11,12,14,15,16)]
		event MobileItem_AttachmentReadEventHandler AttachmentReadEvent;

		/// <summary>
		/// SupportByVersion Outlook 10 11 12 14 15,16
		/// </summary>
		[SupportByVersion("Outlook", 10,11,12,14,15,16)]
		event MobileItem_BeforeAttachmentSaveEventHandler BeforeAttachmentSaveEvent;

		/// <summary>
		/// SupportByVersion Outlook 10 11 12 14 15,16
		/// </summary>
		[SupportByVersion("Outlook", 10,11,12,14,15,16)]
		event MobileItem_BeforeDeleteEventHandler BeforeDeleteEvent;

		/// <summary>
		/// SupportByVersion Outlook 12 14 15,16
		/// </summary>
		[SupportByVersion("Outlook", 12,14,15,16)]
		event MobileItem_AttachmentRemoveEventHandler AttachmentRemoveEvent;

		/// <summary>
		/// SupportByVersion Outlook 12 14 15,16
		/// </summary>
		[SupportByVersion("Outlook", 12,14,15,16)]
		event MobileItem_BeforeAttachmentAddEventHandler BeforeAttachmentAddEvent;

		/// <summary>
		/// SupportByVersion Outlook 12 14 15,16
		/// </summary>
		[SupportByVersion("Outlook", 12,14,15,16)]
		event MobileItem_BeforeAttachmentPreviewEventHandler BeforeAttachmentPreviewEvent;

		/// <summary>
		/// SupportByVersion Outlook 12 14 15,16
		/// </summary>
		[SupportByVersion("Outlook", 12,14,15,16)]
		event MobileItem_BeforeAttachmentReadEventHandler BeforeAttachmentReadEvent;

		/// <summary>
		/// SupportByVersion Outlook 12 14 15,16
		/// </summary>
		[SupportByVersion("Outlook", 12,14,15,16)]
		event MobileItem_BeforeAttachmentWriteToTempFileEventHandler BeforeAttachmentWriteToTempFileEvent;

		/// <summary>
		/// SupportByVersion Outlook 12 14 15,16
		/// </summary>
		[SupportByVersion("Outlook", 12,14,15,16)]
		event MobileItem_UnloadEventHandler UnloadEvent;

		/// <summary>
		/// SupportByVersion Outlook 12 14 15,16
		/// </summary>
		[SupportByVersion("Outlook", 12,14,15,16)]
		event MobileItem_BeforeAutoSaveEventHandler BeforeAutoSaveEvent;

		/// <summary>
		/// SupportByVersion Outlook 14 15,16
		/// </summary>
		[SupportByVersion("Outlook", 14,15,16)]
		event MobileItem_BeforeReadEventHandler BeforeReadEvent;

		/// <summary>
		/// SupportByVersion Outlook 14 15,16
		/// </summary>
		[SupportByVersion("Outlook", 14,15,16)]
		event MobileItem_AfterWriteEventHandler AfterWriteEvent;

		/// <summary>
		/// SupportByVersion Outlook 15,16
		/// </summary>
		[SupportByVersion("Outlook", 15, 16)]
		event MobileItem_ReadCompleteEventHandler ReadCompleteEvent;

        #endregion
    }
}
