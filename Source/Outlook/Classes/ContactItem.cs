using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.OutlookApi
{
	#region Delegates

	#pragma warning disable
	public delegate void ContactItem_OpenEventHandler(ref bool cancel);
	public delegate void ContactItem_CustomActionEventHandler(ICOMObject action, ICOMObject response, ref bool cancel);
	public delegate void ContactItem_CustomPropertyChangeEventHandler(string name);
	public delegate void ContactItem_ForwardEventHandler(ICOMObject forward, ref bool cancel);
	public delegate void ContactItem_CloseEventHandler(ref bool cancel);
	public delegate void ContactItem_PropertyChangeEventHandler(string name);
	public delegate void ContactItem_ReadEventHandler();
	public delegate void ContactItem_ReplyEventHandler(ICOMObject response, ref bool cancel);
	public delegate void ContactItem_ReplyAllEventHandler(ICOMObject response, ref bool cancel);
	public delegate void ContactItem_SendEventHandler(ref bool cancel);
	public delegate void ContactItem_WriteEventHandler(ref bool cancel);
	public delegate void ContactItem_BeforeCheckNamesEventHandler(ref bool cancel);
	public delegate void ContactItem_AttachmentAddEventHandler(NetOffice.OutlookApi.Attachment attachment);
	public delegate void ContactItem_AttachmentReadEventHandler(NetOffice.OutlookApi.Attachment attachment);
	public delegate void ContactItem_BeforeAttachmentSaveEventHandler(NetOffice.OutlookApi.Attachment attachment, ref bool cancel);
	public delegate void ContactItem_BeforeDeleteEventHandler(ICOMObject item, ref bool cancel);
	public delegate void ContactItem_AttachmentRemoveEventHandler(NetOffice.OutlookApi.Attachment attachment);
	public delegate void ContactItem_BeforeAttachmentAddEventHandler(NetOffice.OutlookApi.Attachment attachment, ref bool cancel);
	public delegate void ContactItem_BeforeAttachmentPreviewEventHandler(NetOffice.OutlookApi.Attachment attachment, ref bool cancel);
	public delegate void ContactItem_BeforeAttachmentReadEventHandler(NetOffice.OutlookApi.Attachment attachment, ref bool cancel);
	public delegate void ContactItem_BeforeAttachmentWriteToTempFileEventHandler(NetOffice.OutlookApi.Attachment attachment, ref bool cancel);
	public delegate void ContactItem_UnloadEventHandler();
	public delegate void ContactItem_BeforeAutoSaveEventHandler(ref bool cancel);
	public delegate void ContactItem_BeforeReadEventHandler();
	public delegate void ContactItem_AfterWriteEventHandler();
	public delegate void ContactItem_ReadCompleteEventHandler(ref bool cancel);
	#pragma warning restore

	#endregion

	/// <summary>
	/// CoClass ContactItem 
	/// SupportByVersion Outlook, 9,10,11,12,14,15,16
	/// </summary>
	/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff867603.aspx </remarks>
	[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
	[EntityType(EntityType.IsCoClass)]
    [ComEventContract(typeof(EventContracts.ItemEvents), typeof(EventContracts.ItemEvents_10))]
	[TypeId("00061031-0000-0000-C000-000000000046")]
    public interface ContactItem : _ContactItem, IEventBinding
	{
        #region Events

        /// <summary>
        /// SupportByVersion Outlook 9 10 11 12 14 15,16
        /// </summary>
        ///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff867143.aspx </remarks>
        [SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		event ContactItem_OpenEventHandler OpenEvent;

		/// <summary>
		/// SupportByVersion Outlook 9 10 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff869585.aspx </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		event ContactItem_CustomActionEventHandler CustomActionEvent;

		/// <summary>
		/// SupportByVersion Outlook 9 10 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff864389.aspx </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		event ContactItem_CustomPropertyChangeEventHandler CustomPropertyChangeEvent;

		/// <summary>
		/// SupportByVersion Outlook 9 10 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff869224.aspx </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		event ContactItem_ForwardEventHandler ForwardEvent;

		/// <summary>
		/// SupportByVersion Outlook 9 10 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff868853.aspx </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		event ContactItem_CloseEventHandler CloseEvent;

		/// <summary>
		/// SupportByVersion Outlook 9 10 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff864009.aspx </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		event ContactItem_PropertyChangeEventHandler PropertyChangeEvent;

		/// <summary>
		/// SupportByVersion Outlook 9 10 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff864989.aspx </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		event ContactItem_ReadEventHandler ReadEvent;

		/// <summary>
		/// SupportByVersion Outlook 9 10 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff860448.aspx </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		event ContactItem_ReplyEventHandler ReplyEvent;

		/// <summary>
		/// SupportByVersion Outlook 9 10 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff863599.aspx </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		event ContactItem_ReplyAllEventHandler ReplyAllEvent;

		/// <summary>
		/// SupportByVersion Outlook 9 10 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff862690.aspx </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		event ContactItem_SendEventHandler SendEvent;

		/// <summary>
		/// SupportByVersion Outlook 9 10 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff867819.aspx </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		event ContactItem_WriteEventHandler WriteEvent;

		/// <summary>
		/// SupportByVersion Outlook 9 10 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff866923.aspx </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		event ContactItem_BeforeCheckNamesEventHandler BeforeCheckNamesEvent;

		/// <summary>
		/// SupportByVersion Outlook 9 10 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff869827.aspx </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		event ContactItem_AttachmentAddEventHandler AttachmentAddEvent;

		/// <summary>
		/// SupportByVersion Outlook 9 10 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff865586.aspx </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		event ContactItem_AttachmentReadEventHandler AttachmentReadEvent;

		/// <summary>
		/// SupportByVersion Outlook 9 10 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff868975.aspx </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		event ContactItem_BeforeAttachmentSaveEventHandler BeforeAttachmentSaveEvent;

		/// <summary>
		/// SupportByVersion Outlook 10 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff868306.aspx </remarks>
		[SupportByVersion("Outlook", 10,11,12,14,15,16)]
		event ContactItem_BeforeDeleteEventHandler BeforeDeleteEvent;

		/// <summary>
		/// SupportByVersion Outlook 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff869649.aspx </remarks>
		[SupportByVersion("Outlook", 12,14,15,16)]
		event ContactItem_AttachmentRemoveEventHandler AttachmentRemoveEvent;

		/// <summary>
		/// SupportByVersion Outlook 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff869227.aspx </remarks>
		[SupportByVersion("Outlook", 12,14,15,16)]
		event ContactItem_BeforeAttachmentAddEventHandler BeforeAttachmentAddEvent;

		/// <summary>
		/// SupportByVersion Outlook 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff866588.aspx </remarks>
		[SupportByVersion("Outlook", 12,14,15,16)]
		event ContactItem_BeforeAttachmentPreviewEventHandler BeforeAttachmentPreviewEvent;

		/// <summary>
		/// SupportByVersion Outlook 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff868776.aspx </remarks>
		[SupportByVersion("Outlook", 12,14,15,16)]
		event ContactItem_BeforeAttachmentReadEventHandler BeforeAttachmentReadEvent;

		/// <summary>
		/// SupportByVersion Outlook 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff869347.aspx </remarks>
		[SupportByVersion("Outlook", 12,14,15,16)]
		event ContactItem_BeforeAttachmentWriteToTempFileEventHandler BeforeAttachmentWriteToTempFileEvent;

		/// <summary>
		/// SupportByVersion Outlook 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff861599.aspx </remarks>
		[SupportByVersion("Outlook", 12,14,15,16)]
		event ContactItem_UnloadEventHandler UnloadEvent;

		/// <summary>
		/// SupportByVersion Outlook 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff869089.aspx </remarks>
		[SupportByVersion("Outlook", 12,14,15,16)]
		event ContactItem_BeforeAutoSaveEventHandler BeforeAutoSaveEvent;

		/// <summary>
		/// SupportByVersion Outlook 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff869176.aspx </remarks>
		[SupportByVersion("Outlook", 14,15,16)]
		event ContactItem_BeforeReadEventHandler BeforeReadEvent;

		/// <summary>
		/// SupportByVersion Outlook 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff869360.aspx </remarks>
		[SupportByVersion("Outlook", 14,15,16)]
		event ContactItem_AfterWriteEventHandler AfterWriteEvent;

		/// <summary>
		/// SupportByVersion Outlook 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/jj227683.aspx </remarks>
		[SupportByVersion("Outlook", 15, 16)]
		event ContactItem_ReadCompleteEventHandler ReadCompleteEvent;

        #endregion
    }
}
