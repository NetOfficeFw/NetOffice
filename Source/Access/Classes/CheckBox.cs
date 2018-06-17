using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.AccessApi
{
	#region Delegates

	#pragma warning disable
	public delegate void CheckBox_ClickEventHandler();
	public delegate void CheckBox_BeforeUpdateEventHandler(ref Int16 cancel);
	public delegate void CheckBox_AfterUpdateEventHandler();
	public delegate void CheckBox_EnterEventHandler();
	public delegate void CheckBox_ExitEventHandler(ref Int16 cancel);
	public delegate void CheckBox_GotFocusEventHandler();
	public delegate void CheckBox_LostFocusEventHandler();
	public delegate void CheckBox_DblClickEventHandler(ref Int16 cancel);
	public delegate void CheckBox_MouseDownEventHandler(ref Int16 button, ref Int16 shift, ref Single x, ref Single y);
	public delegate void CheckBox_MouseMoveEventHandler(ref Int16 button, ref Int16 shift, ref Single x, ref Single y);
	public delegate void CheckBox_MouseUpEventHandler(ref Int16 button, ref Int16 shift, ref Single x, ref Single y);
	public delegate void CheckBox_KeyDownEventHandler(ref Int16 keyCode, ref Int16 shift);
	public delegate void CheckBox_KeyPressEventHandler(ref Int16 keyAscii);
	public delegate void CheckBox_KeyUpEventHandler(ref Int16 keyCode, ref Int16 shift);
	#pragma warning restore

	#endregion

	/// <summary>
	/// CoClass CheckBox
	/// SupportByVersion Access, 9,10,11,12,14,15,16
	/// </summary>
	/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194967.aspx </remarks>
	[SupportByVersion("Access", 9,10,11,12,14,15,16)]
	[EntityType(EntityType.IsCoClass)]
    [ComEventContract(typeof(EventContracts._CheckBoxEvents), typeof(EventContracts.DispCheckBoxEvents))]
	[TypeId("3B06E953-E47C-11CD-8701-00AA003F0F07")]
    public interface CheckBox : _Checkbox, IEventBinding
	{
		#region Events

		/// <summary>
		/// SupportByVersion Access 9 10 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff845501.aspx </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		event CheckBox_ClickEventHandler ClickEvent;

		/// <summary>
		/// SupportByVersion Access 9 10 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff834411.aspx </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		event CheckBox_BeforeUpdateEventHandler BeforeUpdateEvent;

		/// <summary>
		/// SupportByVersion Access 9 10 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff835644.aspx </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		event CheckBox_AfterUpdateEventHandler AfterUpdateEvent;

		/// <summary>
		/// SupportByVersion Access 9 10 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff193836.aspx </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		event CheckBox_EnterEventHandler EnterEvent;

		/// <summary>
		/// SupportByVersion Access 9 10 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff194475.aspx </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		event CheckBox_ExitEventHandler ExitEvent;

		/// <summary>
		/// SupportByVersion Access 9 10 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff192446.aspx </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		event CheckBox_GotFocusEventHandler GotFocusEvent;

		/// <summary>
		/// SupportByVersion Access 9 10 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff822462.aspx </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		event CheckBox_LostFocusEventHandler LostFocusEvent;

		/// <summary>
		/// SupportByVersion Access 9 10 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff835425.aspx </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		event CheckBox_DblClickEventHandler DblClickEvent;

		/// <summary>
		/// SupportByVersion Access 9 10 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff194926.aspx </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		event CheckBox_MouseDownEventHandler MouseDownEvent;

		/// <summary>
		/// SupportByVersion Access 9 10 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff836690.aspx </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		event CheckBox_MouseMoveEventHandler MouseMoveEvent;

		/// <summary>
		/// SupportByVersion Access 9 10 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff195727.aspx </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		event CheckBox_MouseUpEventHandler MouseUpEvent;

		/// <summary>
		/// SupportByVersion Access 9 10 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff845602.aspx </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		event CheckBox_KeyDownEventHandler KeyDownEvent;

		/// <summary>
		/// SupportByVersion Access 9 10 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff197676.aspx </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		event CheckBox_KeyPressEventHandler KeyPressEvent;

		/// <summary>
		/// SupportByVersion Access 9 10 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff193646.aspx </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		event CheckBox_KeyUpEventHandler KeyUpEvent;

		#endregion
	}
}
