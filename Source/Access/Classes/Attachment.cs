using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.AccessApi
{
	#region Delegates

	#pragma warning disable
	public delegate void Attachment_BeforeUpdateEventHandler(ref Int16 cancel);
	public delegate void Attachment_AfterUpdateEventHandler();
	public delegate void Attachment_EnterEventHandler();
	public delegate void Attachment_ExitEventHandler(ref Int16 cancel);
	public delegate void Attachment_DirtyEventHandler(ref Int16 cancel);
	public delegate void Attachment_ChangeEventHandler();
	public delegate void Attachment_GotFocusEventHandler();
	public delegate void Attachment_LostFocusEventHandler();
	public delegate void Attachment_ClickEventHandler();
	public delegate void Attachment_DblClickEventHandler(ref Int16 cancel);
	public delegate void Attachment_MouseDownEventHandler(ref Int16 button, ref Int16 shift, ref Single x, ref Single y);
	public delegate void Attachment_MouseMoveEventHandler(ref Int16 button, ref Int16 shift, ref Single x, ref Single y);
	public delegate void Attachment_MouseUpEventHandler(ref Int16 button, ref Int16 shift, ref Single x, ref Single y);
	public delegate void Attachment_KeyDownEventHandler(ref Int16 keyCode, ref Int16 shift);
	public delegate void Attachment_KeyPressEventHandler(ref Int16 keyAscii);
	public delegate void Attachment_KeyUpEventHandler(ref Int16 keyCode, ref Int16 shift);
	public delegate void Attachment_AttachmentCurrentEventHandler();
	#pragma warning restore

	#endregion

	/// <summary>
	/// CoClass Attachment
	/// SupportByVersion Access, 12,14,15,16
	/// </summary>
	/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff821783.aspx </remarks>
	[SupportByVersion("Access", 12,14,15,16)]
	[EntityType(EntityType.IsCoClass)]
    [ComEventContract(typeof(EventContracts.DispAttachmentEvents))]
	[TypeId("3B06E979-E47C-11CD-8701-00AA003F0F07")]
    public interface Attachment : _Attachment, IEventBinding
	{
		#region Events

		/// <summary>
		/// SupportByVersion Access 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff844829.aspx </remarks>
		[SupportByVersion("Access", 12,14,15,16)]
		event Attachment_BeforeUpdateEventHandler BeforeUpdateEvent;

		/// <summary>
		/// SupportByVersion Access 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff845081.aspx </remarks>
		[SupportByVersion("Access", 12,14,15,16)]
		event Attachment_AfterUpdateEventHandler AfterUpdateEvent;

		/// <summary>
		/// SupportByVersion Access 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff845173.aspx </remarks>
		[SupportByVersion("Access", 12,14,15,16)]
		event Attachment_EnterEventHandler EnterEvent;

		/// <summary>
		/// SupportByVersion Access 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff820770.aspx </remarks>
		[SupportByVersion("Access", 12,14,15,16)]
		event Attachment_ExitEventHandler ExitEvent;

		/// <summary>
		/// SupportByVersion Access 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff834764.aspx </remarks>
		[SupportByVersion("Access", 12,14,15,16)]
		event Attachment_DirtyEventHandler DirtyEvent;

		/// <summary>
		/// SupportByVersion Access 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff194528.aspx </remarks>
		[SupportByVersion("Access", 12,14,15,16)]
		event Attachment_ChangeEventHandler ChangeEvent;

		/// <summary>
		/// SupportByVersion Access 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff198117.aspx </remarks>
		[SupportByVersion("Access", 12,14,15,16)]
		event Attachment_GotFocusEventHandler GotFocusEvent;

		/// <summary>
		/// SupportByVersion Access 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff822030.aspx </remarks>
		[SupportByVersion("Access", 12,14,15,16)]
		event Attachment_LostFocusEventHandler LostFocusEvent;

		/// <summary>
		/// SupportByVersion Access 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff834489.aspx </remarks>
		[SupportByVersion("Access", 12,14,15,16)]
		event Attachment_ClickEventHandler ClickEvent;

		/// <summary>
		/// SupportByVersion Access 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff821484.aspx </remarks>
		[SupportByVersion("Access", 12,14,15,16)]
		event Attachment_DblClickEventHandler DblClickEvent;

		/// <summary>
		/// SupportByVersion Access 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff193169.aspx </remarks>
		[SupportByVersion("Access", 12,14,15,16)]
		event Attachment_MouseDownEventHandler MouseDownEvent;

		/// <summary>
		/// SupportByVersion Access 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff194908.aspx </remarks>
		[SupportByVersion("Access", 12,14,15,16)]
		event Attachment_MouseMoveEventHandler MouseMoveEvent;

		/// <summary>
		/// SupportByVersion Access 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff821765.aspx </remarks>
		[SupportByVersion("Access", 12,14,15,16)]
		event Attachment_MouseUpEventHandler MouseUpEvent;

		/// <summary>
		/// SupportByVersion Access 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff197635.aspx </remarks>
		[SupportByVersion("Access", 12,14,15,16)]
		event Attachment_KeyDownEventHandler KeyDownEvent;

		/// <summary>
		/// SupportByVersion Access 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff837202.aspx </remarks>
		[SupportByVersion("Access", 12,14,15,16)]
		event Attachment_KeyPressEventHandler KeyPressEvent;

		/// <summary>
		/// SupportByVersion Access 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff193501.aspx </remarks>
		[SupportByVersion("Access", 12,14,15,16)]
		event Attachment_KeyUpEventHandler KeyUpEvent;

		/// <summary>
		/// SupportByVersion Access 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff193515.aspx </remarks>
		[SupportByVersion("Access", 12,14,15,16)]
		event Attachment_AttachmentCurrentEventHandler AttachmentCurrentEvent;

		#endregion
	}
}
