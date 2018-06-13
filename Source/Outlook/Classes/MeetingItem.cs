using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.OutlookApi
{
	#region Delegates

	#pragma warning disable
	public delegate void MeetingItem_OpenEventHandler(ref bool cancel);
	public delegate void MeetingItem_CustomActionEventHandler(ICOMObject action, ICOMObject response, ref bool cancel);
	public delegate void MeetingItem_CustomPropertyChangeEventHandler(string name);
	public delegate void MeetingItem_ForwardEventHandler(ICOMObject forward, ref bool cancel);
	public delegate void MeetingItem_CloseEventHandler(ref bool cancel);
	public delegate void MeetingItem_PropertyChangeEventHandler(string Name);
	public delegate void MeetingItem_ReadEventHandler();
	public delegate void MeetingItem_ReplyEventHandler(ICOMObject response, ref bool cancel);
	public delegate void MeetingItem_ReplyAllEventHandler(ICOMObject response, ref bool cancel);
	public delegate void MeetingItem_SendEventHandler(ref bool cancel);
	public delegate void MeetingItem_WriteEventHandler(ref bool cancel);
	public delegate void MeetingItem_BeforeCheckNamesEventHandler(ref bool cancel);
	public delegate void MeetingItem_AttachmentAddEventHandler(NetOffice.OutlookApi.Attachment attachment);
	public delegate void MeetingItem_AttachmentReadEventHandler(NetOffice.OutlookApi.Attachment attachment);
	public delegate void MeetingItem_BeforeAttachmentSaveEventHandler(NetOffice.OutlookApi.Attachment attachment, ref bool cancel);
	public delegate void MeetingItem_BeforeDeleteEventHandler(ICOMObject item, ref bool cancel);
	public delegate void MeetingItem_AttachmentRemoveEventHandler(NetOffice.OutlookApi.Attachment attachment);
	public delegate void MeetingItem_BeforeAttachmentAddEventHandler(NetOffice.OutlookApi.Attachment attachment, ref bool cancel);
	public delegate void MeetingItem_BeforeAttachmentPreviewEventHandler(NetOffice.OutlookApi.Attachment attachment, ref bool cancel);
	public delegate void MeetingItem_BeforeAttachmentReadEventHandler(NetOffice.OutlookApi.Attachment attachment, ref bool cancel);
	public delegate void MeetingItem_BeforeAttachmentWriteToTempFileEventHandler(NetOffice.OutlookApi.Attachment attachment, ref bool cancel);
	public delegate void MeetingItem_UnloadEventHandler();
	public delegate void MeetingItem_BeforeAutoSaveEventHandler(ref bool cancel);
	public delegate void MeetingItem_BeforeReadEventHandler();
	public delegate void MeetingItem_AfterWriteEventHandler();
	public delegate void MeetingItem_ReadCompleteEventHandler(ref bool cancel);
	#pragma warning restore

	#endregion

	/// <summary>
	/// CoClass MeetingItem 
	/// SupportByVersion Outlook, 9,10,11,12,14,15,16
	/// </summary>
	/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff868714.aspx </remarks>
	[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
	[EntityType(EntityType.IsCoClass)]
    [ComEventContract(typeof(EventContracts.ItemEvents), typeof(EventContracts.ItemEvents_10))]
	[TypeId("00061036-0000-0000-C000-000000000046")]
    public interface MeetingItem : _MeetingItem, IEventBinding
	{
        #region Events

        /// <summary>
        /// SupportByVersion Outlook 9 10 11 12 14 15,16
        /// </summary>
        ///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff869262.aspx </remarks>
        [SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		event MeetingItem_OpenEventHandler OpenEvent;

		/// <summary>
		/// SupportByVersion Outlook 9 10 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff869085.aspx </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		event MeetingItem_CustomActionEventHandler CustomActionEvent;

		/// <summary>
		/// SupportByVersion Outlook 9 10 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff868648.aspx </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		event MeetingItem_CustomPropertyChangeEventHandler CustomPropertyChangeEvent;

		/// <summary>
		/// SupportByVersion Outlook 9 10 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff860965.aspx </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		event MeetingItem_ForwardEventHandler ForwardEvent;

		/// <summary>
		/// SupportByVersion Outlook 9 10 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff868082.aspx </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		event MeetingItem_CloseEventHandler CloseEvent;

		/// <summary>
		/// SupportByVersion Outlook 9 10 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff866218.aspx </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		event MeetingItem_PropertyChangeEventHandler PropertyChangeEvent;

		/// <summary>
		/// SupportByVersion Outlook 9 10 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff867466.aspx </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		event MeetingItem_ReadEventHandler ReadEvent;

		/// <summary>
		/// SupportByVersion Outlook 9 10 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff865388.aspx </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		event MeetingItem_ReplyEventHandler ReplyEvent;

		/// <summary>
		/// SupportByVersion Outlook 9 10 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff869401.aspx </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		event MeetingItem_ReplyAllEventHandler ReplyAllEvent;

		/// <summary>
		/// SupportByVersion Outlook 9 10 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff868190.aspx </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		event MeetingItem_SendEventHandler SendEvent;

		/// <summary>
		/// SupportByVersion Outlook 9 10 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff862376.aspx </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		event MeetingItem_WriteEventHandler WriteEvent;

		/// <summary>
		/// SupportByVersion Outlook 9 10 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff864280.aspx </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		event MeetingItem_BeforeCheckNamesEventHandler BeforeCheckNamesEvent;

		/// <summary>
		/// SupportByVersion Outlook 9 10 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff869709.aspx </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		event MeetingItem_AttachmentAddEventHandler AttachmentAddEvent;

		/// <summary>
		/// SupportByVersion Outlook 9 10 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff864995.aspx </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		event MeetingItem_AttachmentReadEventHandler AttachmentReadEvent;

		/// <summary>
		/// SupportByVersion Outlook 9 10 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff862132.aspx </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		event MeetingItem_BeforeAttachmentSaveEventHandler BeforeAttachmentSaveEvent;

		/// <summary>
		/// SupportByVersion Outlook 10 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff861564.aspx </remarks>
		[SupportByVersion("Outlook", 10,11,12,14,15,16)]
		event MeetingItem_BeforeDeleteEventHandler BeforeDeleteEvent;

		/// <summary>
		/// SupportByVersion Outlook 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff864184.aspx </remarks>
		[SupportByVersion("Outlook", 12,14,15,16)]
		event MeetingItem_AttachmentRemoveEventHandler AttachmentRemoveEvent;

		/// <summary>
		/// SupportByVersion Outlook 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff867860.aspx </remarks>
		[SupportByVersion("Outlook", 12,14,15,16)]
		event MeetingItem_BeforeAttachmentAddEventHandler BeforeAttachmentAddEvent;

		/// <summary>
		/// SupportByVersion Outlook 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff864699.aspx </remarks>
		[SupportByVersion("Outlook", 12,14,15,16)]
		event MeetingItem_BeforeAttachmentPreviewEventHandler BeforeAttachmentPreviewEvent;

		/// <summary>
		/// SupportByVersion Outlook 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff861629.aspx </remarks>
		[SupportByVersion("Outlook", 12,14,15,16)]
		event MeetingItem_BeforeAttachmentReadEventHandler BeforeAttachmentReadEvent;

		/// <summary>
		/// SupportByVersion Outlook 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff862521.aspx </remarks>
		[SupportByVersion("Outlook", 12,14,15,16)]
		event MeetingItem_BeforeAttachmentWriteToTempFileEventHandler BeforeAttachmentWriteToTempFileEvent;

		/// <summary>
		/// SupportByVersion Outlook 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff867350.aspx </remarks>
		[SupportByVersion("Outlook", 12,14,15,16)]
		event MeetingItem_UnloadEventHandler UnloadEvent;

		/// <summary>
		/// SupportByVersion Outlook 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff865356.aspx </remarks>
		[SupportByVersion("Outlook", 12,14,15,16)]
		event MeetingItem_BeforeAutoSaveEventHandler BeforeAutoSaveEvent;

		/// <summary>
		/// SupportByVersion Outlook 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff869422.aspx </remarks>
		[SupportByVersion("Outlook", 14,15,16)]
		event MeetingItem_BeforeReadEventHandler BeforeReadEvent;

		/// <summary>
		/// SupportByVersion Outlook 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff861265.aspx </remarks>
		[SupportByVersion("Outlook", 14,15,16)]
		event MeetingItem_AfterWriteEventHandler AfterWriteEvent;

		/// <summary>
		/// SupportByVersion Outlook 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/jj227688.aspx </remarks>
		[SupportByVersion("Outlook", 15, 16)]
		event MeetingItem_ReadCompleteEventHandler ReadCompleteEvent;

        #endregion
    }
}
