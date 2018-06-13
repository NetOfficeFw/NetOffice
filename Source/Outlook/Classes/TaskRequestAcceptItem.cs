using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.OutlookApi
{
	#region Delegates

	#pragma warning disable
	public delegate void TaskRequestAcceptItem_OpenEventHandler(ref bool cancel);
	public delegate void TaskRequestAcceptItem_CustomActionEventHandler(ICOMObject action, ICOMObject response, ref bool cancel);
	public delegate void TaskRequestAcceptItem_CustomPropertyChangeEventHandler(string name);
	public delegate void TaskRequestAcceptItem_ForwardEventHandler(ICOMObject forward, ref bool cancel);
	public delegate void TaskRequestAcceptItem_CloseEventHandler(ref bool cancel);
	public delegate void TaskRequestAcceptItem_PropertyChangeEventHandler(string Name);
	public delegate void TaskRequestAcceptItem_ReadEventHandler();
	public delegate void TaskRequestAcceptItem_ReplyEventHandler(ICOMObject response, ref bool cancel);
	public delegate void TaskRequestAcceptItem_ReplyAllEventHandler(ICOMObject response, ref bool cancel);
	public delegate void TaskRequestAcceptItem_SendEventHandler(ref bool cancel);
	public delegate void TaskRequestAcceptItem_WriteEventHandler(ref bool cancel);
	public delegate void TaskRequestAcceptItem_BeforeCheckNamesEventHandler(ref bool cancel);
	public delegate void TaskRequestAcceptItem_AttachmentAddEventHandler(NetOffice.OutlookApi.Attachment attachment);
	public delegate void TaskRequestAcceptItem_AttachmentReadEventHandler(NetOffice.OutlookApi.Attachment attachment);
	public delegate void TaskRequestAcceptItem_BeforeAttachmentSaveEventHandler(NetOffice.OutlookApi.Attachment attachment, ref bool cancel);
	public delegate void TaskRequestAcceptItem_BeforeDeleteEventHandler(ICOMObject item, ref bool cancel);
	public delegate void TaskRequestAcceptItem_AttachmentRemoveEventHandler(NetOffice.OutlookApi.Attachment attachment);
	public delegate void TaskRequestAcceptItem_BeforeAttachmentAddEventHandler(NetOffice.OutlookApi.Attachment attachment, ref bool cancel);
	public delegate void TaskRequestAcceptItem_BeforeAttachmentPreviewEventHandler(NetOffice.OutlookApi.Attachment attachment, ref bool cancel);
	public delegate void TaskRequestAcceptItem_BeforeAttachmentReadEventHandler(NetOffice.OutlookApi.Attachment attachment, ref bool cancel);
	public delegate void TaskRequestAcceptItem_BeforeAttachmentWriteToTempFileEventHandler(NetOffice.OutlookApi.Attachment attachment, ref bool cancel);
	public delegate void TaskRequestAcceptItem_UnloadEventHandler();
	public delegate void TaskRequestAcceptItem_BeforeAutoSaveEventHandler(ref bool cancel);
	public delegate void TaskRequestAcceptItem_BeforeReadEventHandler();
	public delegate void TaskRequestAcceptItem_AfterWriteEventHandler();
	public delegate void TaskRequestAcceptItem_ReadCompleteEventHandler(ref bool cancel);
	#pragma warning restore

	#endregion

	/// <summary>
	/// CoClass TaskRequestAcceptItem 
	/// SupportByVersion Outlook, 9,10,11,12,14,15,16
	/// </summary>
	/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff868287.aspx </remarks>
	[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
	[EntityType(EntityType.IsCoClass)]
    [ComEventContract(typeof(EventContracts.ItemEvents), typeof(EventContracts.ItemEvents_10))]
	[TypeId("00061052-0000-0000-C000-000000000046")]
    public interface TaskRequestAcceptItem : _TaskRequestAcceptItem, IEventBinding
	{
        #region Events

        /// <summary>
        /// SupportByVersion Outlook 9 10 11 12 14 15,16
        /// </summary>
        ///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff864475.aspx </remarks>
        [SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		event TaskRequestAcceptItem_OpenEventHandler OpenEvent;

		/// <summary>
		/// SupportByVersion Outlook 9 10 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff866911.aspx </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		event TaskRequestAcceptItem_CustomActionEventHandler CustomActionEvent;

		/// <summary>
		/// SupportByVersion Outlook 9 10 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff865677.aspx </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		event TaskRequestAcceptItem_CustomPropertyChangeEventHandler CustomPropertyChangeEvent;

		/// <summary>
		/// SupportByVersion Outlook 9 10 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff864251.aspx </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		event TaskRequestAcceptItem_ForwardEventHandler ForwardEvent;

		/// <summary>
		/// SupportByVersion Outlook 9 10 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff867821.aspx </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		event TaskRequestAcceptItem_CloseEventHandler CloseEvent;

		/// <summary>
		/// SupportByVersion Outlook 9 10 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff864697.aspx </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		event TaskRequestAcceptItem_PropertyChangeEventHandler PropertyChangeEvent;

		/// <summary>
		/// SupportByVersion Outlook 9 10 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff862748.aspx </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		event TaskRequestAcceptItem_ReadEventHandler ReadEvent;

		/// <summary>
		/// SupportByVersion Outlook 9 10 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff869840.aspx </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		event TaskRequestAcceptItem_ReplyEventHandler ReplyEvent;

		/// <summary>
		/// SupportByVersion Outlook 9 10 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff863705.aspx </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		event TaskRequestAcceptItem_ReplyAllEventHandler ReplyAllEvent;

		/// <summary>
		/// SupportByVersion Outlook 9 10 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff864417.aspx </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		event TaskRequestAcceptItem_SendEventHandler SendEvent;

		/// <summary>
		/// SupportByVersion Outlook 9 10 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff860301.aspx </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		event TaskRequestAcceptItem_WriteEventHandler WriteEvent;

		/// <summary>
		/// SupportByVersion Outlook 9 10 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff861561.aspx </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		event TaskRequestAcceptItem_BeforeCheckNamesEventHandler BeforeCheckNamesEvent;

		/// <summary>
		/// SupportByVersion Outlook 9 10 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff869987.aspx </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		event TaskRequestAcceptItem_AttachmentAddEventHandler AttachmentAddEvent;

		/// <summary>
		/// SupportByVersion Outlook 9 10 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff863003.aspx </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		event TaskRequestAcceptItem_AttachmentReadEventHandler AttachmentReadEvent;

		/// <summary>
		/// SupportByVersion Outlook 9 10 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff866757.aspx </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		event TaskRequestAcceptItem_BeforeAttachmentSaveEventHandler BeforeAttachmentSaveEvent;

		/// <summary>
		/// SupportByVersion Outlook 10 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff867104.aspx </remarks>
		[SupportByVersion("Outlook", 10,11,12,14,15,16)]
		event TaskRequestAcceptItem_BeforeDeleteEventHandler BeforeDeleteEvent;

		/// <summary>
		/// SupportByVersion Outlook 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff865598.aspx </remarks>
		[SupportByVersion("Outlook", 12,14,15,16)]
		event TaskRequestAcceptItem_AttachmentRemoveEventHandler AttachmentRemoveEvent;

		/// <summary>
		/// SupportByVersion Outlook 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff867237.aspx </remarks>
		[SupportByVersion("Outlook", 12,14,15,16)]
		event TaskRequestAcceptItem_BeforeAttachmentAddEventHandler BeforeAttachmentAddEvent;

		/// <summary>
		/// SupportByVersion Outlook 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff868108.aspx </remarks>
		[SupportByVersion("Outlook", 12,14,15,16)]
		event TaskRequestAcceptItem_BeforeAttachmentPreviewEventHandler BeforeAttachmentPreviewEvent;

		/// <summary>
		/// SupportByVersion Outlook 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff865264.aspx </remarks>
		[SupportByVersion("Outlook", 12,14,15,16)]
		event TaskRequestAcceptItem_BeforeAttachmentReadEventHandler BeforeAttachmentReadEvent;

		/// <summary>
		/// SupportByVersion Outlook 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff865985.aspx </remarks>
		[SupportByVersion("Outlook", 12,14,15,16)]
		event TaskRequestAcceptItem_BeforeAttachmentWriteToTempFileEventHandler BeforeAttachmentWriteToTempFileEvent;

		/// <summary>
		/// SupportByVersion Outlook 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff861852.aspx </remarks>
		[SupportByVersion("Outlook", 12,14,15,16)]
		event TaskRequestAcceptItem_UnloadEventHandler UnloadEvent;

		/// <summary>
		/// SupportByVersion Outlook 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff860393.aspx </remarks>
		[SupportByVersion("Outlook", 12,14,15,16)]
		event TaskRequestAcceptItem_BeforeAutoSaveEventHandler BeforeAutoSaveEvent;

		/// <summary>
		/// SupportByVersion Outlook 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff866727.aspx </remarks>
		[SupportByVersion("Outlook", 14,15,16)]
		event TaskRequestAcceptItem_BeforeReadEventHandler BeforeReadEvent;

		/// <summary>
		/// SupportByVersion Outlook 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff869882.aspx </remarks>
		[SupportByVersion("Outlook", 14,15,16)]
		event TaskRequestAcceptItem_AfterWriteEventHandler AfterWriteEvent;

		/// <summary>
		/// SupportByVersion Outlook 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/jj230145.aspx </remarks>
		[SupportByVersion("Outlook", 15, 16)]
		event TaskRequestAcceptItem_ReadCompleteEventHandler ReadCompleteEvent;

        #endregion
    }
}
