using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.AccessApi
{
	#region Delegates

	#pragma warning disable
	public delegate void TextBox_BeforeUpdateEventHandler(ref Int16 cancel);
	public delegate void TextBox_AfterUpdateEventHandler();
	public delegate void TextBox_ChangeEventHandler();
	public delegate void TextBox_EnterEventHandler();
	public delegate void TextBox_ExitEventHandler(ref Int16 cancel);
	public delegate void TextBox_GotFocusEventHandler();
	public delegate void TextBox_LostFocusEventHandler();
	public delegate void TextBox_ClickEventHandler();
	public delegate void TextBox_DblClickEventHandler(ref Int16 cancel);
	public delegate void TextBox_MouseDownEventHandler(ref Int16 button, ref Int16 shift, ref Single x, ref Single y);
	public delegate void TextBox_MouseMoveEventHandler(ref Int16 button, ref Int16 shift, ref Single x, ref Single y);
	public delegate void TextBox_MouseUpEventHandler(ref Int16 button, ref Int16 shift, ref Single x, ref Single y);
	public delegate void TextBox_KeyDownEventHandler(ref Int16 keyCode, ref Int16 shift);
	public delegate void TextBox_KeyPressEventHandler(ref Int16 keyAscii);
	public delegate void TextBox_KeyUpEventHandler(ref Int16 keyCode, ref Int16 shift);
	public delegate void TextBox_DirtyEventHandler(ref Int16 cancel);
	public delegate void TextBox_UndoEventHandler(ref Int16 cancel);
	#pragma warning restore

	#endregion

	/// <summary>
	/// CoClass TextBox 
	/// SupportByVersion Access, 9,10,11,12,14,15,16
	/// </summary>
	/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff835063.aspx </remarks>
	[SupportByVersion("Access", 9,10,11,12,14,15,16)]
	[EntityType(EntityType.IsCoClass)]
    [ComEventContract(typeof(EventContracts._TextBoxEvents), typeof(EventContracts.DispTextBoxEvents))]
	[TypeId("3B06E945-E47C-11CD-8701-00AA003F0F07")]
    public interface TextBox : _Textbox, IEventBinding
	{
		#region Events

		/// <summary>
		/// SupportByVersion Access 9 10 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff845199.aspx </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		event TextBox_BeforeUpdateEventHandler BeforeUpdateEvent;

		/// <summary>
		/// SupportByVersion Access 9 10 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff194818.aspx </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		event TextBox_AfterUpdateEventHandler AfterUpdateEvent;

		/// <summary>
		/// SupportByVersion Access 9 10 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff821734.aspx </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		event TextBox_ChangeEventHandler ChangeEvent;

		/// <summary>
		/// SupportByVersion Access 9 10 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff197769.aspx </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		event TextBox_EnterEventHandler EnterEvent;

		/// <summary>
		/// SupportByVersion Access 9 10 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff844925.aspx </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		event TextBox_ExitEventHandler ExitEvent;

		/// <summary>
		/// SupportByVersion Access 9 10 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff822716.aspx </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		event TextBox_GotFocusEventHandler GotFocusEvent;

		/// <summary>
		/// SupportByVersion Access 9 10 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff193542.aspx </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		event TextBox_LostFocusEventHandler LostFocusEvent;

		/// <summary>
		/// SupportByVersion Access 9 10 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff834731.aspx </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		event TextBox_ClickEventHandler ClickEvent;

		/// <summary>
		/// SupportByVersion Access 9 10 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff821748.aspx </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		event TextBox_DblClickEventHandler DblClickEvent;

		/// <summary>
		/// SupportByVersion Access 9 10 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff821739.aspx </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		event TextBox_MouseDownEventHandler MouseDownEvent;

		/// <summary>
		/// SupportByVersion Access 9 10 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff197411.aspx </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		event TextBox_MouseMoveEventHandler MouseMoveEvent;

		/// <summary>
		/// SupportByVersion Access 9 10 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff845232.aspx </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		event TextBox_MouseUpEventHandler MouseUpEvent;

		/// <summary>
		/// SupportByVersion Access 9 10 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff844722.aspx </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		event TextBox_KeyDownEventHandler KeyDownEvent;

		/// <summary>
		/// SupportByVersion Access 9 10 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff197040.aspx </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		event TextBox_KeyPressEventHandler KeyPressEvent;

		/// <summary>
		/// SupportByVersion Access 9 10 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff191709.aspx </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		event TextBox_KeyUpEventHandler KeyUpEvent;

		/// <summary>
		/// SupportByVersion Access 10 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff835038.aspx </remarks>
		[SupportByVersion("Access", 10,11,12,14,15,16)]
		event TextBox_DirtyEventHandler DirtyEvent;

		/// <summary>
		/// SupportByVersion Access 10 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff836364.aspx </remarks>
		[SupportByVersion("Access", 10,11,12,14,15,16)]
		event TextBox_UndoEventHandler UndoEvent;

		#endregion
	}
}
