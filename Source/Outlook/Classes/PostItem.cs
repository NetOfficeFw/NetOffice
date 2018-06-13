using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.OutlookApi
{
	#region Delegates

	#pragma warning disable
	public delegate void PostItem_OpenEventHandler(ref bool cancel);
	public delegate void PostItem_CustomActionEventHandler(ICOMObject action, ICOMObject response, ref bool cancel);
	public delegate void PostItem_CustomPropertyChangeEventHandler(string name);
	public delegate void PostItem_ForwardEventHandler(ICOMObject forward, ref bool cancel);
	public delegate void PostItem_CloseEventHandler(ref bool cancel);
	public delegate void PostItem_PropertyChangeEventHandler(string name);
	public delegate void PostItem_ReadEventHandler();
	public delegate void PostItem_ReplyEventHandler(ICOMObject response, ref bool cancel);
	public delegate void PostItem_ReplyAllEventHandler(ICOMObject response, ref bool cancel);
	public delegate void PostItem_SendEventHandler(ref bool cancel);
	public delegate void PostItem_WriteEventHandler(ref bool cancel);
	public delegate void PostItem_BeforeCheckNamesEventHandler(ref bool cancel);
	public delegate void PostItem_AttachmentAddEventHandler(NetOffice.OutlookApi.Attachment attachment);
	public delegate void PostItem_AttachmentReadEventHandler(NetOffice.OutlookApi.Attachment attachment);
	public delegate void PostItem_BeforeAttachmentSaveEventHandler(NetOffice.OutlookApi.Attachment attachment, ref bool cancel);
	public delegate void PostItem_BeforeDeleteEventHandler(ICOMObject item, ref bool cancel);
	public delegate void PostItem_AttachmentRemoveEventHandler(NetOffice.OutlookApi.Attachment attachment);
	public delegate void PostItem_BeforeAttachmentAddEventHandler(NetOffice.OutlookApi.Attachment attachment, ref bool cancel);
	public delegate void PostItem_BeforeAttachmentPreviewEventHandler(NetOffice.OutlookApi.Attachment attachment, ref bool cancel);
	public delegate void PostItem_BeforeAttachmentReadEventHandler(NetOffice.OutlookApi.Attachment attachment, ref bool cancel);
	public delegate void PostItem_BeforeAttachmentWriteToTempFileEventHandler(NetOffice.OutlookApi.Attachment attachment, ref bool cancel);
	public delegate void PostItem_UnloadEventHandler();
	public delegate void PostItem_BeforeAutoSaveEventHandler(ref bool cancel);
	public delegate void PostItem_BeforeReadEventHandler();
	public delegate void PostItem_AfterWriteEventHandler();
	public delegate void PostItem_ReadCompleteEventHandler(ref bool cancel);
	#pragma warning restore

	#endregion

	/// <summary>
	/// CoClass PostItem 
	/// SupportByVersion Outlook, 9,10,11,12,14,15,16
	/// </summary>
	/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff869495.aspx </remarks>
	[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
	[EntityType(EntityType.IsCoClass)]
    [ComEventContract(typeof(EventContracts.ItemEvents), typeof(EventContracts.ItemEvents_10))]
	[TypeId("0006103A-0000-0000-C000-000000000046")]
    public interface PostItem : _PostItem, IEventBinding
	{
        #region Events

        /// <summary>
        /// SupportByVersion Outlook 9 10 11 12 14 15,16
        /// </summary>
        ///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff868576.aspx </remarks>
        [SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		event PostItem_OpenEventHandler OpenEvent;

		/// <summary>
		/// SupportByVersion Outlook 9 10 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff865988.aspx </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		event PostItem_CustomActionEventHandler CustomActionEvent;

		/// <summary>
		/// SupportByVersion Outlook 9 10 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff869699.aspx </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		event PostItem_CustomPropertyChangeEventHandler CustomPropertyChangeEvent;

		/// <summary>
		/// SupportByVersion Outlook 9 10 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff869634.aspx </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		event PostItem_ForwardEventHandler ForwardEvent;

		/// <summary>
		/// SupportByVersion Outlook 9 10 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff865788.aspx </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		event PostItem_CloseEventHandler CloseEvent;

		/// <summary>
		/// SupportByVersion Outlook 9 10 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff866428.aspx </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		event PostItem_PropertyChangeEventHandler PropertyChangeEvent;

		/// <summary>
		/// SupportByVersion Outlook 9 10 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff863971.aspx </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		event PostItem_ReadEventHandler ReadEvent;

		/// <summary>
		/// SupportByVersion Outlook 9 10 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff864006.aspx </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		event PostItem_ReplyEventHandler ReplyEvent;

		/// <summary>
		/// SupportByVersion Outlook 9 10 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff864196.aspx </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		event PostItem_ReplyAllEventHandler ReplyAllEvent;

		/// <summary>
		/// SupportByVersion Outlook 9 10 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff869234.aspx </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		event PostItem_SendEventHandler SendEvent;

		/// <summary>
		/// SupportByVersion Outlook 9 10 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff862670.aspx </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		event PostItem_WriteEventHandler WriteEvent;

		/// <summary>
		/// SupportByVersion Outlook 9 10 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff868684.aspx </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		event PostItem_BeforeCheckNamesEventHandler BeforeCheckNamesEvent;

		/// <summary>
		/// SupportByVersion Outlook 9 10 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff867873.aspx </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		event PostItem_AttachmentAddEventHandler AttachmentAddEvent;

		/// <summary>
		/// SupportByVersion Outlook 9 10 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff863919.aspx </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		event PostItem_AttachmentReadEventHandler AttachmentReadEvent;

		/// <summary>
		/// SupportByVersion Outlook 9 10 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff865082.aspx </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		event PostItem_BeforeAttachmentSaveEventHandler BeforeAttachmentSaveEvent;

		/// <summary>
		/// SupportByVersion Outlook 10 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff868959.aspx </remarks>
		[SupportByVersion("Outlook", 10,11,12,14,15,16)]
		event PostItem_BeforeDeleteEventHandler BeforeDeleteEvent;

		/// <summary>
		/// SupportByVersion Outlook 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff868227.aspx </remarks>
		[SupportByVersion("Outlook", 12,14,15,16)]
		event PostItem_AttachmentRemoveEventHandler AttachmentRemoveEvent;

		/// <summary>
		/// SupportByVersion Outlook 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff865099.aspx </remarks>
		[SupportByVersion("Outlook", 12,14,15,16)]
		event PostItem_BeforeAttachmentAddEventHandler BeforeAttachmentAddEvent;

		/// <summary>
		/// SupportByVersion Outlook 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff861933.aspx </remarks>
		[SupportByVersion("Outlook", 12,14,15,16)]
		event PostItem_BeforeAttachmentPreviewEventHandler BeforeAttachmentPreviewEvent;

		/// <summary>
		/// SupportByVersion Outlook 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff868979.aspx </remarks>
		[SupportByVersion("Outlook", 12,14,15,16)]
		event PostItem_BeforeAttachmentReadEventHandler BeforeAttachmentReadEvent;

		/// <summary>
		/// SupportByVersion Outlook 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff868891.aspx </remarks>
		[SupportByVersion("Outlook", 12,14,15,16)]
		event PostItem_BeforeAttachmentWriteToTempFileEventHandler BeforeAttachmentWriteToTempFileEvent;

		/// <summary>
		/// SupportByVersion Outlook 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff864217.aspx </remarks>
		[SupportByVersion("Outlook", 12,14,15,16)]
		event PostItem_UnloadEventHandler UnloadEvent;

		/// <summary>
		/// SupportByVersion Outlook 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff865806.aspx </remarks>
		[SupportByVersion("Outlook", 12,14,15,16)]
		event PostItem_BeforeAutoSaveEventHandler BeforeAutoSaveEvent;

		/// <summary>
		/// SupportByVersion Outlook 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff862520.aspx </remarks>
		[SupportByVersion("Outlook", 14,15,16)]
		event PostItem_BeforeReadEventHandler BeforeReadEvent;

		/// <summary>
		/// SupportByVersion Outlook 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff869570.aspx </remarks>
		[SupportByVersion("Outlook", 14,15,16)]
		event PostItem_AfterWriteEventHandler AfterWriteEvent;

		/// <summary>
		/// SupportByVersion Outlook 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/jj229578.aspx </remarks>
		[SupportByVersion("Outlook", 15, 16)]
		event PostItem_ReadCompleteEventHandler ReadCompleteEvent;

        #endregion
    }
}
