using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.AccessApi
{
	#region Delegates

	#pragma warning disable
	public delegate void ListBox_BeforeUpdateEventHandler(ref Int16 cancel);
	public delegate void ListBox_AfterUpdateEventHandler();
	public delegate void ListBox_EnterEventHandler();
	public delegate void ListBox_ExitEventHandler(ref Int16 cancel);
	public delegate void ListBox_GotFocusEventHandler();
	public delegate void ListBox_LostFocusEventHandler();
	public delegate void ListBox_ClickEventHandler();
	public delegate void ListBox_DblClickEventHandler(ref Int16 cancel);
	public delegate void ListBox_MouseDownEventHandler(ref Int16 button, ref Int16 shift, ref Single x, ref Single y);
	public delegate void ListBox_MouseMoveEventHandler(ref Int16 button, ref Int16 shift, ref Single x, ref Single y);
	public delegate void ListBox_MouseUpEventHandler(ref Int16 button, ref Int16 shift, ref Single x, ref Single y);
	public delegate void ListBox_KeyDownEventHandler(ref Int16 keyCode, ref Int16 shift);
	public delegate void ListBox_KeyPressEventHandler(ref Int16 keyAscii);
	public delegate void ListBox_KeyUpEventHandler(ref Int16 keyCode, ref Int16 shift);
	#pragma warning restore

	#endregion

	/// <summary>
	/// CoClass ListBox 
	/// SupportByVersion Access, 9,10,11,12,14,15,16
	/// </summary>
	/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff195480.aspx </remarks>
	[SupportByVersion("Access", 9,10,11,12,14,15,16)]
	[EntityType(EntityType.IsCoClass)]
    [ComEventContract(typeof(EventContracts._ListBoxEvents), typeof(EventContracts.DispListBoxEvents))]
	[TypeId("3B06E959-E47C-11CD-8701-00AA003F0F07")]
    public interface ListBox : _ListBox, IEventBinding
	{
		#region Events

		/// <summary>
		/// SupportByVersion Access 9 10 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff192062.aspx </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		event ListBox_BeforeUpdateEventHandler BeforeUpdateEvent;

		/// <summary>
		/// SupportByVersion Access 9 10 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff822464.aspx </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		event ListBox_AfterUpdateEventHandler AfterUpdateEvent;

		/// <summary>
		/// SupportByVersion Access 9 10 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff194346.aspx </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		event ListBox_EnterEventHandler EnterEvent;

		/// <summary>
		/// SupportByVersion Access 9 10 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff195415.aspx </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		event ListBox_ExitEventHandler ExitEvent;

		/// <summary>
		/// SupportByVersion Access 9 10 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff822062.aspx </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		event ListBox_GotFocusEventHandler GotFocusEvent;

		/// <summary>
		/// SupportByVersion Access 9 10 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff844967.aspx </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		event ListBox_LostFocusEventHandler LostFocusEvent;

		/// <summary>
		/// SupportByVersion Access 9 10 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff197659.aspx </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		event ListBox_ClickEventHandler ClickEvent;

		/// <summary>
		/// SupportByVersion Access 9 10 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff837260.aspx </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		event ListBox_DblClickEventHandler DblClickEvent;

		/// <summary>
		/// SupportByVersion Access 9 10 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff822715.aspx </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		event ListBox_MouseDownEventHandler MouseDownEvent;

		/// <summary>
		/// SupportByVersion Access 9 10 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff836717.aspx </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		event ListBox_MouseMoveEventHandler MouseMoveEvent;

		/// <summary>
		/// SupportByVersion Access 9 10 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff197349.aspx </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		event ListBox_MouseUpEventHandler MouseUpEvent;

		/// <summary>
		/// SupportByVersion Access 9 10 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff194748.aspx </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		event ListBox_KeyDownEventHandler KeyDownEvent;

		/// <summary>
		/// SupportByVersion Access 9 10 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff845356.aspx </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		event ListBox_KeyPressEventHandler KeyPressEvent;

		/// <summary>
		/// SupportByVersion Access 9 10 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff192236.aspx </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		event ListBox_KeyUpEventHandler KeyUpEvent;

		#endregion
	}
}
