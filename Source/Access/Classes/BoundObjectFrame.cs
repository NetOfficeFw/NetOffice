using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.AccessApi
{
	#region Delegates

	#pragma warning disable
	public delegate void BoundObjectFrame_UpdatedEventHandler(ref Int16 code);
	public delegate void BoundObjectFrame_BeforeUpdateEventHandler(ref Int16 cancel);
	public delegate void BoundObjectFrame_AfterUpdateEventHandler();
	public delegate void BoundObjectFrame_EnterEventHandler();
	public delegate void BoundObjectFrame_ExitEventHandler(ref Int16 cancel);
	public delegate void BoundObjectFrame_GotFocusEventHandler();
	public delegate void BoundObjectFrame_LostFocusEventHandler();
	public delegate void BoundObjectFrame_ClickEventHandler();
	public delegate void BoundObjectFrame_DblClickEventHandler(ref Int16 cancel);
	public delegate void BoundObjectFrame_MouseDownEventHandler(ref Int16 button, ref Int16 shift, ref Single x, ref Single y);
	public delegate void BoundObjectFrame_MouseMoveEventHandler(ref Int16 button, ref Int16 shift, ref Single x, ref Single y);
	public delegate void BoundObjectFrame_MouseUpEventHandler(ref Int16 button, ref Int16 shift, ref Single x, ref Single y);
	public delegate void BoundObjectFrame_KeyDownEventHandler(ref Int16 keyCode, ref Int16 shift);
	public delegate void BoundObjectFrame_KeyPressEventHandler(ref Int16 keyAscii);
	public delegate void BoundObjectFrame_KeyUpEventHandler(ref Int16 keyCode, ref Int16 shift);
	#pragma warning restore

	#endregion

	/// <summary>
	/// CoClass BoundObjectFrame
	/// SupportByVersion Access, 9,10,11,12,14,15,16
	/// </summary>
	/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff822036.aspx </remarks>
	[SupportByVersion("Access", 9,10,11,12,14,15,16)]
	[EntityType(EntityType.IsCoClass)]
    [ComEventContract(typeof(EventContracts._BoundObjectFrameEvents), typeof(EventContracts.DispBoundObjectFrameEvents))]
	[TypeId("3B06E957-E47C-11CD-8701-00AA003F0F07")]
    public interface BoundObjectFrame : _BoundObjectFrame, IEventBinding
	{
		#region Events

		/// <summary>
		/// SupportByVersion Access 9 10 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff836253.aspx </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		event BoundObjectFrame_UpdatedEventHandler UpdatedEvent;

		/// <summary>
		/// SupportByVersion Access 9 10 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff193585.aspx </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		event BoundObjectFrame_BeforeUpdateEventHandler BeforeUpdateEvent;

		/// <summary>
		/// SupportByVersion Access 9 10 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff837257.aspx </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		event BoundObjectFrame_AfterUpdateEventHandler AfterUpdateEvent;

		/// <summary>
		/// SupportByVersion Access 9 10 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff821752.aspx </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		event BoundObjectFrame_EnterEventHandler EnterEvent;

		/// <summary>
		/// SupportByVersion Access 9 10 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff196744.aspx </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		event BoundObjectFrame_ExitEventHandler ExitEvent;

		/// <summary>
		/// SupportByVersion Access 9 10 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff836584.aspx </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		event BoundObjectFrame_GotFocusEventHandler GotFocusEvent;

		/// <summary>
		/// SupportByVersion Access 9 10 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff192331.aspx </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		event BoundObjectFrame_LostFocusEventHandler LostFocusEvent;

		/// <summary>
		/// SupportByVersion Access 9 10 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff845134.aspx </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		event BoundObjectFrame_ClickEventHandler ClickEvent;

		/// <summary>
		/// SupportByVersion Access 9 10 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff192445.aspx </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		event BoundObjectFrame_DblClickEventHandler DblClickEvent;

		/// <summary>
		/// SupportByVersion Access 9 10 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff822859.aspx </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		event BoundObjectFrame_MouseDownEventHandler MouseDownEvent;

		/// <summary>
		/// SupportByVersion Access 9 10 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff194904.aspx </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		event BoundObjectFrame_MouseMoveEventHandler MouseMoveEvent;

		/// <summary>
		/// SupportByVersion Access 9 10 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff834792.aspx </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		event BoundObjectFrame_MouseUpEventHandler MouseUpEvent;

		/// <summary>
		/// SupportByVersion Access 9 10 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff197931.aspx </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		event BoundObjectFrame_KeyDownEventHandler KeyDownEvent;

		/// <summary>
		/// SupportByVersion Access 9 10 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff823123.aspx </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		event BoundObjectFrame_KeyPressEventHandler KeyPressEvent;

		/// <summary>
		/// SupportByVersion Access 9 10 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff845734.aspx </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		event BoundObjectFrame_KeyUpEventHandler KeyUpEvent;

		#endregion
	}
}
