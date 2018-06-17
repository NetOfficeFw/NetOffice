using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.AccessApi
{
	#region Delegates

	#pragma warning disable
	public delegate void _CheckBoxInOption_GotFocusEventHandler();
	public delegate void _CheckBoxInOption_LostFocusEventHandler();
	public delegate void _CheckBoxInOption_MouseDownEventHandler(ref Int16 button, ref Int16 shift, ref Single x, ref Single y);
	public delegate void _CheckBoxInOption_MouseMoveEventHandler(ref Int16 button, ref Int16 shift, ref Single x, ref Single y);
	public delegate void _CheckBoxInOption_MouseUpEventHandler(ref Int16 button, ref Int16 shift, ref Single x, ref Single y);
	public delegate void _CheckBoxInOption_KeyDownEventHandler(ref Int16 keyCode, ref Int16 shift);
	public delegate void _CheckBoxInOption_KeyPressEventHandler(ref Int16 keyAscii);
	public delegate void _CheckBoxInOption_KeyUpEventHandler(ref Int16 keyCode, ref Int16 shift);
	public delegate void _CheckBoxInOption_ClickEventHandler();
	public delegate void _CheckBoxInOption_BeforeUpdateEventHandler(ref Int16 cancel);
	public delegate void _CheckBoxInOption_AfterUpdateEventHandler();
	public delegate void _CheckBoxInOption_EnterEventHandler();
	public delegate void _CheckBoxInOption_ExitEventHandler(ref Int16 cancel);
	public delegate void _CheckBoxInOption_DblClickEventHandler(ref Int16 cancel);
	#pragma warning restore

	#endregion

	/// <summary>
	/// CoClass _CheckBoxInOption
	/// SupportByVersion Access, 9,10,11,12,14,15,16
	/// </summary>
	[SupportByVersion("Access", 9,10,11,12,14,15,16)]
	[EntityType(EntityType.IsCoClass)]
    [ComEventContract(typeof(EventContracts._CheckBoxInOptionEvents), typeof(EventContracts.DispCheckBoxEvents))]
	[TypeId("BC9E435C-F037-11CD-8701-00AA003F0F07")]
    public interface _CheckBoxInOption : _Checkbox, IEventBinding
	{
		#region Events

		/// <summary>
		/// SupportByVersion Access 9 10 11 12 14 15,16
		/// </summary>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		event _CheckBoxInOption_GotFocusEventHandler GotFocusEvent;

		/// <summary>
		/// SupportByVersion Access 9 10 11 12 14 15,16
		/// </summary>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		event _CheckBoxInOption_LostFocusEventHandler LostFocusEvent;

		/// <summary>
		/// SupportByVersion Access 9 10 11 12 14 15,16
		/// </summary>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		event _CheckBoxInOption_MouseDownEventHandler MouseDownEvent;

		/// <summary>
		/// SupportByVersion Access 9 10 11 12 14 15,16
		/// </summary>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		event _CheckBoxInOption_MouseMoveEventHandler MouseMoveEvent;

		/// <summary>
		/// SupportByVersion Access 9 10 11 12 14 15,16
		/// </summary>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		event _CheckBoxInOption_MouseUpEventHandler MouseUpEvent;

		/// <summary>
		/// SupportByVersion Access 9 10 11 12 14 15,16
		/// </summary>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		event _CheckBoxInOption_KeyDownEventHandler KeyDownEvent;

		/// <summary>
		/// SupportByVersion Access 9 10 11 12 14 15,16
		/// </summary>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		event _CheckBoxInOption_KeyPressEventHandler KeyPressEvent;

		/// <summary>
		/// SupportByVersion Access 9 10 11 12 14 15,16
		/// </summary>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		event _CheckBoxInOption_KeyUpEventHandler KeyUpEvent;

		/// <summary>
		/// SupportByVersion Access 12 14 15,16
		/// </summary>
		[SupportByVersion("Access", 12,14,15,16)]
		event _CheckBoxInOption_ClickEventHandler ClickEvent;

		/// <summary>
		/// SupportByVersion Access 12 14 15,16
		/// </summary>
		[SupportByVersion("Access", 12,14,15,16)]
		event _CheckBoxInOption_BeforeUpdateEventHandler BeforeUpdateEvent;

		/// <summary>
		/// SupportByVersion Access 12 14 15,16
		/// </summary>
		[SupportByVersion("Access", 12,14,15,16)]
		event _CheckBoxInOption_AfterUpdateEventHandler AfterUpdateEvent;

		/// <summary>
		/// SupportByVersion Access 12 14 15,16
		/// </summary>
		[SupportByVersion("Access", 12,14,15,16)]
		event _CheckBoxInOption_EnterEventHandler EnterEvent;

		/// <summary>
		/// SupportByVersion Access 12 14 15,16
		/// </summary>
		[SupportByVersion("Access", 12,14,15,16)]
		event _CheckBoxInOption_ExitEventHandler ExitEvent;

		/// <summary>
		/// SupportByVersion Access 12 14 15,16
		/// </summary>
		[SupportByVersion("Access", 12,14,15,16)]
		event _CheckBoxInOption_DblClickEventHandler DblClickEvent;

		#endregion
	}
}
