using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.AccessApi
{
	#region Delegates

	#pragma warning disable
	public delegate void OptionButton_ClickEventHandler();
	public delegate void OptionButton_BeforeUpdateEventHandler(ref Int16 cancel);
	public delegate void OptionButton_AfterUpdateEventHandler();
	public delegate void OptionButton_EnterEventHandler();
	public delegate void OptionButton_ExitEventHandler(ref Int16 cancel);
	public delegate void OptionButton_GotFocusEventHandler();
	public delegate void OptionButton_LostFocusEventHandler();
	public delegate void OptionButton_DblClickEventHandler(ref Int16 cancel);
	public delegate void OptionButton_MouseDownEventHandler(ref Int16 button, ref Int16 shift, ref Single x, ref Single y);
	public delegate void OptionButton_MouseMoveEventHandler(ref Int16 button, ref Int16 shift, ref Single x, ref Single y);
	public delegate void OptionButton_MouseUpEventHandler(ref Int16 button, ref Int16 shift, ref Single x, ref Single y);
	public delegate void OptionButton_KeyDownEventHandler(ref Int16 keyCode, ref Int16 shift);
	public delegate void OptionButton_KeyPressEventHandler(ref Int16 keyAscii);
	public delegate void OptionButton_KeyUpEventHandler(ref Int16 keyCode, ref Int16 shift);
	#pragma warning restore

	#endregion

	/// <summary>
	/// CoClass OptionButton 
	/// SupportByVersion Access, 9,10,11,12,14,15,16
	/// </summary>
	/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff195195.aspx </remarks>
	[SupportByVersion("Access", 9,10,11,12,14,15,16)]
	[EntityType(EntityType.IsCoClass)]
    [ComEventContract(typeof(EventContracts._OptionButtonEvents), typeof(EventContracts.DispOptionButtonEvents))]
	[TypeId("3B06E951-E47C-11CD-8701-00AA003F0F07")]
    public interface OptionButton : _OptionButton, IEventBinding
	{
		#region Events

		/// <summary>
		/// SupportByVersion Access 9 10 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff197959.aspx </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		event OptionButton_ClickEventHandler ClickEvent;

		/// <summary>
		/// SupportByVersion Access 9 10 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff198120.aspx </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		event OptionButton_BeforeUpdateEventHandler BeforeUpdateEvent;

		/// <summary>
		/// SupportByVersion Access 9 10 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff835351.aspx </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		event OptionButton_AfterUpdateEventHandler AfterUpdateEvent;

		/// <summary>
		/// SupportByVersion Access 9 10 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff194932.aspx </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		event OptionButton_EnterEventHandler EnterEvent;

		/// <summary>
		/// SupportByVersion Access 9 10 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff192088.aspx </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		event OptionButton_ExitEventHandler ExitEvent;

		/// <summary>
		/// SupportByVersion Access 9 10 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff836556.aspx </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		event OptionButton_GotFocusEventHandler GotFocusEvent;

		/// <summary>
		/// SupportByVersion Access 9 10 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff836038.aspx </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		event OptionButton_LostFocusEventHandler LostFocusEvent;

		/// <summary>
		/// SupportByVersion Access 9 10 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff192874.aspx </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		event OptionButton_DblClickEventHandler DblClickEvent;

		/// <summary>
		/// SupportByVersion Access 9 10 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff194854.aspx </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		event OptionButton_MouseDownEventHandler MouseDownEvent;

		/// <summary>
		/// SupportByVersion Access 9 10 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff192937.aspx </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		event OptionButton_MouseMoveEventHandler MouseMoveEvent;

		/// <summary>
		/// SupportByVersion Access 9 10 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff194184.aspx </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		event OptionButton_MouseUpEventHandler MouseUpEvent;

		/// <summary>
		/// SupportByVersion Access 9 10 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff197970.aspx </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		event OptionButton_KeyDownEventHandler KeyDownEvent;

		/// <summary>
		/// SupportByVersion Access 9 10 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff835723.aspx </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		event OptionButton_KeyPressEventHandler KeyPressEvent;

		/// <summary>
		/// SupportByVersion Access 9 10 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff194221.aspx </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		event OptionButton_KeyUpEventHandler KeyUpEvent;

		#endregion
	}
}
