using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.OutlookApi
{
	#region Delegates

	#pragma warning disable
	public delegate void RemoteItem_OpenEventHandler(ref bool cancel);
	public delegate void RemoteItem_CustomActionEventHandler(ICOMObject action, ICOMObject response, ref bool cancel);
	public delegate void RemoteItem_CustomPropertyChangeEventHandler(string name);
	public delegate void RemoteItem_ForwardEventHandler(ICOMObject forward, ref bool cancel);
	public delegate void RemoteItem_CloseEventHandler(ref bool cancel);
	public delegate void RemoteItem_PropertyChangeEventHandler(string Name);
	public delegate void RemoteItem_ReadEventHandler();
	public delegate void RemoteItem_ReplyEventHandler(ICOMObject response, ref bool cancel);
	public delegate void RemoteItem_ReplyAllEventHandler(ICOMObject response, ref bool cancel);
	public delegate void RemoteItem_SendEventHandler(ref bool cancel);
	public delegate void RemoteItem_WriteEventHandler(ref bool cancel);
	public delegate void RemoteItem_BeforeCheckNamesEventHandler(ref bool cancel);
	public delegate void RemoteItem_AttachmentAddEventHandler(NetOffice.OutlookApi.Attachment attachment);
	public delegate void RemoteItem_AttachmentReadEventHandler(NetOffice.OutlookApi.Attachment attachment);
	public delegate void RemoteItem_BeforeAttachmentSaveEventHandler(NetOffice.OutlookApi.Attachment attachment, ref bool cancel);
	public delegate void RemoteItem_BeforeDeleteEventHandler(ICOMObject item, ref bool cancel);
	public delegate void RemoteItem_AttachmentRemoveEventHandler(NetOffice.OutlookApi.Attachment attachment);
	public delegate void RemoteItem_BeforeAttachmentAddEventHandler(NetOffice.OutlookApi.Attachment attachment, ref bool cancel);
	public delegate void RemoteItem_BeforeAttachmentPreviewEventHandler(NetOffice.OutlookApi.Attachment attachment, ref bool cancel);
	public delegate void RemoteItem_BeforeAttachmentReadEventHandler(NetOffice.OutlookApi.Attachment attachment, ref bool cancel);
	public delegate void RemoteItem_BeforeAttachmentWriteToTempFileEventHandler(NetOffice.OutlookApi.Attachment attachment, ref bool cancel);
	public delegate void RemoteItem_UnloadEventHandler();
	public delegate void RemoteItem_BeforeAutoSaveEventHandler(ref bool cancel);
	public delegate void RemoteItem_BeforeReadEventHandler();
	public delegate void RemoteItem_AfterWriteEventHandler();
	public delegate void RemoteItem_ReadCompleteEventHandler(ref bool cancel);
	#pragma warning restore

	#endregion

	/// <summary>
	/// CoClass RemoteItem 
	/// SupportByVersion Outlook, 9,10,11,12,14,15,16
	/// </summary>
	/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff865832.aspx </remarks>
	[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
	[EntityType(EntityType.IsCoClass)]
    [ComEventContract(typeof(EventContracts.ItemEvents), typeof(EventContracts.ItemEvents_10))]
	[TypeId("00061060-0000-0000-C000-000000000046")]
    public interface RemoteItem : _RemoteItem, IEventBinding
	{
        #region Events

        /// <summary>
        /// SupportByVersion Outlook 9 10 11 12 14 15,16
        /// </summary>
        ///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff865290.aspx </remarks>
        [SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		event RemoteItem_OpenEventHandler OpenEvent;

		/// <summary>
		/// SupportByVersion Outlook 9 10 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff864722.aspx </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		event RemoteItem_CustomActionEventHandler CustomActionEvent;

		/// <summary>
		/// SupportByVersion Outlook 9 10 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff866479.aspx </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		event RemoteItem_CustomPropertyChangeEventHandler CustomPropertyChangeEvent;

		/// <summary>
		/// SupportByVersion Outlook 9 10 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff869946.aspx </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		event RemoteItem_ForwardEventHandler ForwardEvent;

		/// <summary>
		/// SupportByVersion Outlook 9 10 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff866747.aspx </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		event RemoteItem_CloseEventHandler CloseEvent;

		/// <summary>
		/// SupportByVersion Outlook 9 10 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff865835.aspx </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		event RemoteItem_PropertyChangeEventHandler PropertyChangeEvent;

		/// <summary>
		/// SupportByVersion Outlook 9 10 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff866776.aspx </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		event RemoteItem_ReadEventHandler ReadEvent;

		/// <summary>
		/// SupportByVersion Outlook 9 10 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff864413.aspx </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		event RemoteItem_ReplyEventHandler ReplyEvent;

		/// <summary>
		/// SupportByVersion Outlook 9 10 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff866005.aspx </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		event RemoteItem_ReplyAllEventHandler ReplyAllEvent;

		/// <summary>
		/// SupportByVersion Outlook 9 10 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff866206.aspx </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		event RemoteItem_SendEventHandler SendEvent;

		/// <summary>
		/// SupportByVersion Outlook 9 10 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff868308.aspx </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		event RemoteItem_WriteEventHandler WriteEvent;

		/// <summary>
		/// SupportByVersion Outlook 9 10 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff868637.aspx </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		event RemoteItem_BeforeCheckNamesEventHandler BeforeCheckNamesEvent;

		/// <summary>
		/// SupportByVersion Outlook 9 10 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff866964.aspx </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		event RemoteItem_AttachmentAddEventHandler AttachmentAddEvent;

		/// <summary>
		/// SupportByVersion Outlook 9 10 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff861874.aspx </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		event RemoteItem_AttachmentReadEventHandler AttachmentReadEvent;

		/// <summary>
		/// SupportByVersion Outlook 9 10 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff868796.aspx </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		event RemoteItem_BeforeAttachmentSaveEventHandler BeforeAttachmentSaveEvent;

		/// <summary>
		/// SupportByVersion Outlook 10 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff861037.aspx </remarks>
		[SupportByVersion("Outlook", 10,11,12,14,15,16)]
		event RemoteItem_BeforeDeleteEventHandler BeforeDeleteEvent;

		/// <summary>
		/// SupportByVersion Outlook 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff868633.aspx </remarks>
		[SupportByVersion("Outlook", 12,14,15,16)]
		event RemoteItem_AttachmentRemoveEventHandler AttachmentRemoveEvent;

		/// <summary>
		/// SupportByVersion Outlook 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff860390.aspx </remarks>
		[SupportByVersion("Outlook", 12,14,15,16)]
		event RemoteItem_BeforeAttachmentAddEventHandler BeforeAttachmentAddEvent;

		/// <summary>
		/// SupportByVersion Outlook 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff870146.aspx </remarks>
		[SupportByVersion("Outlook", 12,14,15,16)]
		event RemoteItem_BeforeAttachmentPreviewEventHandler BeforeAttachmentPreviewEvent;

		/// <summary>
		/// SupportByVersion Outlook 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff866476.aspx </remarks>
		[SupportByVersion("Outlook", 12,14,15,16)]
		event RemoteItem_BeforeAttachmentReadEventHandler BeforeAttachmentReadEvent;

		/// <summary>
		/// SupportByVersion Outlook 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff870112.aspx </remarks>
		[SupportByVersion("Outlook", 12,14,15,16)]
		event RemoteItem_BeforeAttachmentWriteToTempFileEventHandler BeforeAttachmentWriteToTempFileEvent;

		/// <summary>
		/// SupportByVersion Outlook 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff867576.aspx </remarks>
		[SupportByVersion("Outlook", 12,14,15,16)]
		event RemoteItem_UnloadEventHandler UnloadEvent;

		/// <summary>
		/// SupportByVersion Outlook 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff869916.aspx </remarks>
		[SupportByVersion("Outlook", 12,14,15,16)]
		event RemoteItem_BeforeAutoSaveEventHandler BeforeAutoSaveEvent;

		/// <summary>
		/// SupportByVersion Outlook 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff868451.aspx </remarks>
		[SupportByVersion("Outlook", 14,15,16)]
		event RemoteItem_BeforeReadEventHandler BeforeReadEvent;

		/// <summary>
		/// SupportByVersion Outlook 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff867134.aspx </remarks>
		[SupportByVersion("Outlook", 14,15,16)]
		event RemoteItem_AfterWriteEventHandler AfterWriteEvent;

		/// <summary>
		/// SupportByVersion Outlook 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/jj227878.aspx </remarks>
		[SupportByVersion("Outlook", 15, 16)]
		event RemoteItem_ReadCompleteEventHandler ReadCompleteEvent;

        #endregion
    }
}
