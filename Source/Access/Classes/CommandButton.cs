using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.AccessApi
{
	#region Delegates

	#pragma warning disable
	public delegate void CommandButton_ClickEventHandler();
	public delegate void CommandButton_EnterEventHandler();
	public delegate void CommandButton_ExitEventHandler(ref Int16 cancel);
	public delegate void CommandButton_GotFocusEventHandler();
	public delegate void CommandButton_LostFocusEventHandler();
	public delegate void CommandButton_DblClickEventHandler(ref Int16 cancel);
	public delegate void CommandButton_MouseDownEventHandler(ref Int16 button, ref Int16 shift, ref Single x, ref Single y);
	public delegate void CommandButton_MouseMoveEventHandler(ref Int16 button, ref Int16 shift, ref Single x, ref Single y);
	public delegate void CommandButton_MouseUpEventHandler(ref Int16 button, ref Int16 shift, ref Single x, ref Single y);
	public delegate void CommandButton_KeyDownEventHandler(ref Int16 keyCode, ref Int16 shift);
	public delegate void CommandButton_KeyPressEventHandler(ref Int16 keyAscii);
	public delegate void CommandButton_KeyUpEventHandler(ref Int16 keyCode, ref Int16 shift);
	#pragma warning restore

	#endregion

	/// <summary>
	/// CoClass CommandButton 
	/// SupportByVersion Access, 9,10,11,12,14,15,16
	/// </summary>
	/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff191876.aspx </remarks>
	[SupportByVersion("Access", 9,10,11,12,14,15,16)]
	[EntityType(EntityType.IsCoClass)]
    [ComEventContract(typeof(EventContracts._CommandButtonEvents), typeof(EventContracts.DispCommandButtonEvents))]
	[TypeId("3B06E94F-E47C-11CD-8701-00AA003F0F07")]
    public interface CommandButton : _CommandButton, IEventBinding
	{
		#region Events

		/// <summary>
		/// SupportByVersion Access 9 10 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff822439.aspx </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		event CommandButton_ClickEventHandler ClickEvent;

		/// <summary>
		/// SupportByVersion Access 9 10 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff834424.aspx </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		event CommandButton_EnterEventHandler EnterEvent;

		/// <summary>
		/// SupportByVersion Access 9 10 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff834787.aspx </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		event CommandButton_ExitEventHandler ExitEvent;

		/// <summary>
		/// SupportByVersion Access 9 10 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff822451.aspx </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		event CommandButton_GotFocusEventHandler GotFocusEvent;

		/// <summary>
		/// SupportByVersion Access 9 10 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff821407.aspx </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		event CommandButton_LostFocusEventHandler LostFocusEvent;

		/// <summary>
		/// SupportByVersion Access 9 10 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff845152.aspx </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		event CommandButton_DblClickEventHandler DblClickEvent;

		/// <summary>
		/// SupportByVersion Access 9 10 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff197358.aspx </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		event CommandButton_MouseDownEventHandler MouseDownEvent;

		/// <summary>
		/// SupportByVersion Access 9 10 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff836610.aspx </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		event CommandButton_MouseMoveEventHandler MouseMoveEvent;

		/// <summary>
		/// SupportByVersion Access 9 10 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff197657.aspx </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		event CommandButton_MouseUpEventHandler MouseUpEvent;

		/// <summary>
		/// SupportByVersion Access 9 10 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff834776.aspx </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		event CommandButton_KeyDownEventHandler KeyDownEvent;

		/// <summary>
		/// SupportByVersion Access 9 10 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff821769.aspx </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		event CommandButton_KeyPressEventHandler KeyPressEvent;

		/// <summary>
		/// SupportByVersion Access 9 10 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff195077.aspx </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		event CommandButton_KeyUpEventHandler KeyUpEvent;

		#endregion
	}
}
