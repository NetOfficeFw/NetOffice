using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.AccessApi
{
	#region Delegates

	#pragma warning disable
	public delegate void _ToggleButtonInOption_GotFocusEventHandler();
	public delegate void _ToggleButtonInOption_LostFocusEventHandler();
	public delegate void _ToggleButtonInOption_MouseDownEventHandler(ref Int16 button, ref Int16 shift, ref Single x, ref Single y);
	public delegate void _ToggleButtonInOption_MouseMoveEventHandler(ref Int16 button, ref Int16 shift, ref Single x, ref Single y);
	public delegate void _ToggleButtonInOption_MouseUpEventHandler(ref Int16 button, ref Int16 shift, ref Single x, ref Single y);
	public delegate void _ToggleButtonInOption_KeyDownEventHandler(ref Int16 keyCode, ref Int16 shift);
	public delegate void _ToggleButtonInOption_KeyPressEventHandler(ref Int16 keyAscii);
	public delegate void _ToggleButtonInOption_KeyUpEventHandler(ref Int16 keyCode, ref Int16 shift);
	public delegate void _ToggleButtonInOption_ClickEventHandler();
	public delegate void _ToggleButtonInOption_BeforeUpdateEventHandler(ref Int16 cancel);
	public delegate void _ToggleButtonInOption_AfterUpdateEventHandler();
	public delegate void _ToggleButtonInOption_EnterEventHandler();
	public delegate void _ToggleButtonInOption_ExitEventHandler(ref Int16 cancel);
	public delegate void _ToggleButtonInOption_DblClickEventHandler(ref Int16 cancel);
	#pragma warning restore

	#endregion

	/// <summary>
	/// CoClass _ToggleButtonInOption
	/// SupportByVersion Access, 9,10,11,12,14,15,16
	/// </summary>
	[SupportByVersion("Access", 9,10,11,12,14,15,16)]
	[EntityType(EntityType.IsCoClass)]
    [ComEventContract(typeof(EventContracts._ToggleButtonInOptionEvents), typeof(EventContracts.DispToggleButtonEvents))]
	[TypeId("BC9E435E-F037-11CD-8701-00AA003F0F07")]
    public interface _ToggleButtonInOption : _ToggleButton, IEventBinding
	{
		#region Events

		/// <summary>
		/// SupportByVersion Access 9 10 11 12 14 15,16
		/// </summary>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
        event _ToggleButtonInOption_GotFocusEventHandler GotFocusEvent;

		/// <summary>
		/// SupportByVersion Access 9 10 11 12 14 15,16
		/// </summary>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		event _ToggleButtonInOption_LostFocusEventHandler LostFocusEvent;

		/// <summary>
		/// SupportByVersion Access 9 10 11 12 14 15,16
		/// </summary>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		event _ToggleButtonInOption_MouseDownEventHandler MouseDownEvent;

		/// <summary>
		/// SupportByVersion Access 9 10 11 12 14 15,16
		/// </summary>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		event _ToggleButtonInOption_MouseMoveEventHandler MouseMoveEvent;

		/// <summary>
		/// SupportByVersion Access 9 10 11 12 14 15,16
		/// </summary>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		event _ToggleButtonInOption_MouseUpEventHandler MouseUpEvent;

		/// <summary>
		/// SupportByVersion Access 9 10 11 12 14 15,16
		/// </summary>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		event _ToggleButtonInOption_KeyDownEventHandler KeyDownEvent;

		/// <summary>
		/// SupportByVersion Access 9 10 11 12 14 15,16
		/// </summary>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		event _ToggleButtonInOption_KeyPressEventHandler KeyPressEvent;

		/// <summary>
		/// SupportByVersion Access 9 10 11 12 14 15,16
		/// </summary>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		event _ToggleButtonInOption_KeyUpEventHandler KeyUpEvent;

		/// <summary>
		/// SupportByVersion Access 12 14 15,16
		/// </summary>
		[SupportByVersion("Access", 12,14,15,16)]
		event _ToggleButtonInOption_ClickEventHandler ClickEvent;

		/// <summary>
		/// SupportByVersion Access 12 14 15,16
		/// </summary>
		[SupportByVersion("Access", 12,14,15,16)]
		event _ToggleButtonInOption_BeforeUpdateEventHandler BeforeUpdateEvent;

		/// <summary>
		/// SupportByVersion Access 12 14 15,16
		/// </summary>
		[SupportByVersion("Access", 12,14,15,16)]
		event _ToggleButtonInOption_AfterUpdateEventHandler AfterUpdateEvent;

		/// <summary>
		/// SupportByVersion Access 12 14 15,16
		/// </summary>
		[SupportByVersion("Access", 12,14,15,16)]
		event _ToggleButtonInOption_EnterEventHandler EnterEvent;

		/// <summary>
		/// SupportByVersion Access 12 14 15,16
		/// </summary>
		[SupportByVersion("Access", 12,14,15,16)]
		event _ToggleButtonInOption_ExitEventHandler ExitEvent;

		/// <summary>
		/// SupportByVersion Access 12 14 15,16
		/// </summary>
		[SupportByVersion("Access", 12,14,15,16)]
		event _ToggleButtonInOption_DblClickEventHandler DblClickEvent;

		#endregion
	}
}
