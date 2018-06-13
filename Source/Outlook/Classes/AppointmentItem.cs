using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.OutlookApi
{
	#region Delegates

	#pragma warning disable
	public delegate void AppointmentItem_OpenEventHandler(ref bool cancel);
	public delegate void AppointmentItem_CustomActionEventHandler(ICOMObject action, ICOMObject response, ref bool cancel);
	public delegate void AppointmentItem_CustomPropertyChangeEventHandler(string name);
	public delegate void AppointmentItem_ForwardEventHandler(ICOMObject forward, ref bool cancel);
	public delegate void AppointmentItem_CloseEventHandler(ref bool cancel);
	public delegate void AppointmentItem_PropertyChangeEventHandler(string Name);
	public delegate void AppointmentItem_ReadEventHandler();
	public delegate void AppointmentItem_ReplyEventHandler(ICOMObject response, ref bool cancel);
	public delegate void AppointmentItem_ReplyAllEventHandler(COMObject response, ref bool cancel);
	public delegate void AppointmentItem_SendEventHandler(ref bool cancel);
	public delegate void AppointmentItem_WriteEventHandler(ref bool cancel);
	public delegate void AppointmentItem_BeforeCheckNamesEventHandler(ref bool cancel);
	public delegate void AppointmentItem_AttachmentAddEventHandler(NetOffice.OutlookApi.Attachment attachment);
	public delegate void AppointmentItem_AttachmentReadEventHandler(NetOffice.OutlookApi.Attachment attachment);
	public delegate void AppointmentItem_BeforeAttachmentSaveEventHandler(NetOffice.OutlookApi.Attachment attachment, ref bool cancel);
	public delegate void AppointmentItem_BeforeDeleteEventHandler(ICOMObject item, ref bool cancel);
	public delegate void AppointmentItem_AttachmentRemoveEventHandler(NetOffice.OutlookApi.Attachment attachment);
	public delegate void AppointmentItem_BeforeAttachmentAddEventHandler(NetOffice.OutlookApi.Attachment attachment, ref bool cancel);
	public delegate void AppointmentItem_BeforeAttachmentPreviewEventHandler(NetOffice.OutlookApi.Attachment attachment, ref bool cancel);
	public delegate void AppointmentItem_BeforeAttachmentReadEventHandler(NetOffice.OutlookApi.Attachment attachment, ref bool cancel);
	public delegate void AppointmentItem_BeforeAttachmentWriteToTempFileEventHandler(NetOffice.OutlookApi.Attachment attachment, ref bool cancel);
	public delegate void AppointmentItem_UnloadEventHandler();
	public delegate void AppointmentItem_BeforeAutoSaveEventHandler(ref bool cancel);
	public delegate void AppointmentItem_BeforeReadEventHandler();
	public delegate void AppointmentItem_AfterWriteEventHandler();
	public delegate void AppointmentItem_ReadCompleteEventHandler(ref bool cancel);
	#pragma warning restore

	#endregion

	/// <summary>
	/// CoClass AppointmentItem 
	/// SupportByVersion Outlook, 9,10,11,12,14,15,16
	/// </summary>
	/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff862177.aspx </remarks>
	[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
	[EntityType(EntityType.IsCoClass)]
	[TypeId("00061030-0000-0000-C000-000000000046")]
	public interface AppointmentItem : _AppointmentItem, IEventBinding
	{
        #region Events

        /// <summary>
        /// SupportByVersion Outlook 9 10 11 12 14 15,16
        /// </summary>
        ///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff860682.aspx </remarks>
        [SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		event AppointmentItem_OpenEventHandler OpenEvent;

		/// <summary>
		/// SupportByVersion Outlook 9 10 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff868820.aspx </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		event AppointmentItem_CustomActionEventHandler CustomActionEvent;

		/// <summary>
		/// SupportByVersion Outlook 9 10 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff869927.aspx </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		event AppointmentItem_CustomPropertyChangeEventHandler CustomPropertyChangeEvent;

		/// <summary>
		/// SupportByVersion Outlook 9 10 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff863899.aspx </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		event AppointmentItem_ForwardEventHandler ForwardEvent;

		/// <summary>
		/// SupportByVersion Outlook 9 10 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff861899.aspx </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		event AppointmentItem_CloseEventHandler CloseEvent;

		/// <summary>
		/// SupportByVersion Outlook 9 10 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff867180.aspx </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		event AppointmentItem_PropertyChangeEventHandler PropertyChangeEvent;

		/// <summary>
		/// SupportByVersion Outlook 9 10 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff868450.aspx </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		event AppointmentItem_ReadEventHandler ReadEvent;

		/// <summary>
		/// SupportByVersion Outlook 9 10 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff868805.aspx </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		event AppointmentItem_ReplyEventHandler ReplyEvent;

		/// <summary>
		/// SupportByVersion Outlook 9 10 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff868971.aspx </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		event AppointmentItem_ReplyAllEventHandler ReplyAllEvent;

		/// <summary>
		/// SupportByVersion Outlook 9 10 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff865990.aspx </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		event AppointmentItem_SendEventHandler SendEvent;

		/// <summary>
		/// SupportByVersion Outlook 9 10 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff865085.aspx </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		event AppointmentItem_WriteEventHandler WriteEvent;

		/// <summary>
		/// SupportByVersion Outlook 9 10 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff869640.aspx </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		event AppointmentItem_BeforeCheckNamesEventHandler BeforeCheckNamesEvent;

		/// <summary>
		/// SupportByVersion Outlook 9 10 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff864491.aspx </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		event AppointmentItem_AttachmentAddEventHandler AttachmentAddEvent;

		/// <summary>
		/// SupportByVersion Outlook 9 10 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff866077.aspx </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		event AppointmentItem_AttachmentReadEventHandler AttachmentReadEvent;

		/// <summary>
		/// SupportByVersion Outlook 9 10 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff861892.aspx </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		event AppointmentItem_BeforeAttachmentSaveEventHandler BeforeAttachmentSaveEvent;

		/// <summary>
		/// SupportByVersion Outlook 10 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff869458.aspx </remarks>
		[SupportByVersion("Outlook", 10,11,12,14,15,16)]
		event AppointmentItem_BeforeDeleteEventHandler BeforeDeleteEvent;

		/// <summary>
		/// SupportByVersion Outlook 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff869436.aspx </remarks>
		[SupportByVersion("Outlook", 12,14,15,16)]
		event AppointmentItem_AttachmentRemoveEventHandler AttachmentRemoveEvent;

		/// <summary>
		/// SupportByVersion Outlook 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff861568.aspx </remarks>
		[SupportByVersion("Outlook", 12,14,15,16)]
		event AppointmentItem_BeforeAttachmentAddEventHandler BeforeAttachmentAddEvent;

		/// <summary>
		/// SupportByVersion Outlook 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff867460.aspx </remarks>
		[SupportByVersion("Outlook", 12,14,15,16)]
		event AppointmentItem_BeforeAttachmentPreviewEventHandler BeforeAttachmentPreviewEvent;

		/// <summary>
		/// SupportByVersion Outlook 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff869437.aspx </remarks>
		[SupportByVersion("Outlook", 12,14,15,16)]
		event AppointmentItem_BeforeAttachmentReadEventHandler BeforeAttachmentReadEvent;

		/// <summary>
		/// SupportByVersion Outlook 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff866752.aspx </remarks>
		[SupportByVersion("Outlook", 12,14,15,16)]
		event AppointmentItem_BeforeAttachmentWriteToTempFileEventHandler BeforeAttachmentWriteToTempFileEvent;

		/// <summary>
		/// SupportByVersion Outlook 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff867872.aspx </remarks>
		[SupportByVersion("Outlook", 12,14,15,16)]
		event AppointmentItem_UnloadEventHandler UnloadEvent;

		/// <summary>
		/// SupportByVersion Outlook 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff868926.aspx </remarks>
		[SupportByVersion("Outlook", 12,14,15,16)]
		event AppointmentItem_BeforeAutoSaveEventHandler BeforeAutoSaveEvent;

		/// <summary>
		/// SupportByVersion Outlook 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff868991.aspx </remarks>
		[SupportByVersion("Outlook", 14,15,16)]
		event AppointmentItem_BeforeReadEventHandler BeforeReadEvent;

		/// <summary>
		/// SupportByVersion Outlook 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff869638.aspx </remarks>
		[SupportByVersion("Outlook", 14,15,16)]
		event AppointmentItem_AfterWriteEventHandler AfterWriteEvent;

		/// <summary>
		/// SupportByVersion Outlook 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/jj229473.aspx </remarks>
		[SupportByVersion("Outlook", 15, 16)]
		event AppointmentItem_ReadCompleteEventHandler ReadCompleteEvent;

        #endregion
    }
}
