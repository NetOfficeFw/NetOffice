using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.AccessApi
{
	#region Delegates

	#pragma warning disable
	public delegate void ToggleButton_ClickEventHandler();
	public delegate void ToggleButton_BeforeUpdateEventHandler(ref Int16 cancel);
	public delegate void ToggleButton_AfterUpdateEventHandler();
	public delegate void ToggleButton_EnterEventHandler();
	public delegate void ToggleButton_ExitEventHandler(ref Int16 cancel);
	public delegate void ToggleButton_GotFocusEventHandler();
	public delegate void ToggleButton_LostFocusEventHandler();
	public delegate void ToggleButton_DblClickEventHandler(ref Int16 cancel);
	public delegate void ToggleButton_MouseDownEventHandler(ref Int16 button, ref Int16 shift, ref Single x, ref Single y);
	public delegate void ToggleButton_MouseMoveEventHandler(ref Int16 button, ref Int16 shift, ref Single x, ref Single y);
	public delegate void ToggleButton_MouseUpEventHandler(ref Int16 button, ref Int16 shift, ref Single x, ref Single y);
	public delegate void ToggleButton_KeyDownEventHandler(ref Int16 keyCode, ref Int16 shift);
	public delegate void ToggleButton_KeyPressEventHandler(ref Int16 keyAscii);
	public delegate void ToggleButton_KeyUpEventHandler(ref Int16 keyCode, ref Int16 shift);
	#pragma warning restore

	#endregion

	/// <summary>
	/// CoClass ToggleButton 
	/// SupportByVersion Access, 9,10,11,12,14,15,16
	/// </summary>
	/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff845729.aspx </remarks>
	[SupportByVersion("Access", 9,10,11,12,14,15,16)]
	[EntityType(EntityType.IsCoClass)]
    [ComEventContract(typeof(EventContracts._ToggleButtonEvents), typeof(EventContracts.DispToggleButtonEvents))]
	[TypeId("3B06E961-E47C-11CD-8701-00AA003F0F07")]
    public interface ToggleButton : _ToggleButton, IEventBinding
	{
		#region Events

		/// <summary>
		/// SupportByVersion Access 9 10 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff822512.aspx </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		event ToggleButton_ClickEventHandler ClickEvent;

		/// <summary>
		/// SupportByVersion Access 9 10 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff193550.aspx </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		event ToggleButton_BeforeUpdateEventHandler BeforeUpdateEvent;

		/// <summary>
		/// SupportByVersion Access 9 10 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff197370.aspx </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		event ToggleButton_AfterUpdateEventHandler AfterUpdateEvent;

		/// <summary>
		/// SupportByVersion Access 9 10 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff822076.aspx </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		event ToggleButton_EnterEventHandler EnterEvent;

		/// <summary>
		/// SupportByVersion Access 9 10 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff822741.aspx </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		event ToggleButton_ExitEventHandler ExitEvent;

		/// <summary>
		/// SupportByVersion Access 9 10 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff844947.aspx </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		event ToggleButton_GotFocusEventHandler GotFocusEvent;

		/// <summary>
		/// SupportByVersion Access 9 10 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff192720.aspx </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		event ToggleButton_LostFocusEventHandler LostFocusEvent;

		/// <summary>
		/// SupportByVersion Access 9 10 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff835061.aspx </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		event ToggleButton_DblClickEventHandler DblClickEvent;

		/// <summary>
		/// SupportByVersion Access 9 10 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff193526.aspx </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		event ToggleButton_MouseDownEventHandler MouseDownEvent;

		/// <summary>
		/// SupportByVersion Access 9 10 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff821745.aspx </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		event ToggleButton_MouseMoveEventHandler MouseMoveEvent;

		/// <summary>
		/// SupportByVersion Access 9 10 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff821377.aspx </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		event ToggleButton_MouseUpEventHandler MouseUpEvent;

		/// <summary>
		/// SupportByVersion Access 9 10 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff195763.aspx </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		event ToggleButton_KeyDownEventHandler KeyDownEvent;

		/// <summary>
		/// SupportByVersion Access 9 10 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff192251.aspx </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		event ToggleButton_KeyPressEventHandler KeyPressEvent;

		/// <summary>
		/// SupportByVersion Access 9 10 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff197674.aspx </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		event ToggleButton_KeyUpEventHandler KeyUpEvent;

		#endregion
	}
}
