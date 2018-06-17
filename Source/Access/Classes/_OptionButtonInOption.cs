using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.AccessApi
{
	#region Delegates

	#pragma warning disable
	public delegate void _OptionButtonInOption_GotFocusEventHandler();
	public delegate void _OptionButtonInOption_LostFocusEventHandler();
	public delegate void _OptionButtonInOption_MouseDownEventHandler(ref Int16 button, ref Int16 shift, ref Single x, ref Single y);
	public delegate void _OptionButtonInOption_MouseMoveEventHandler(ref Int16 button, ref Int16 shift, ref Single x, ref Single y);
	public delegate void _OptionButtonInOption_MouseUpEventHandler(ref Int16 button, ref Int16 shift, ref Single x, ref Single y);
	public delegate void _OptionButtonInOption_KeyDownEventHandler(ref Int16 keyCode, ref Int16 shift);
	public delegate void _OptionButtonInOption_KeyPressEventHandler(ref Int16 keyAscii);
	public delegate void _OptionButtonInOption_KeyUpEventHandler(ref Int16 keyCode, ref Int16 shift);
	public delegate void _OptionButtonInOption_ClickEventHandler();
	public delegate void _OptionButtonInOption_BeforeUpdateEventHandler(ref Int16 cancel);
	public delegate void _OptionButtonInOption_AfterUpdateEventHandler();
	public delegate void _OptionButtonInOption_EnterEventHandler();
	public delegate void _OptionButtonInOption_ExitEventHandler(ref Int16 cancel);
	public delegate void _OptionButtonInOption_DblClickEventHandler(ref Int16 cancel);
	#pragma warning restore

	#endregion

	/// <summary>
	/// CoClass _OptionButtonInOption
	/// SupportByVersion Access, 9,10,11,12,14,15,16
	/// </summary>
	[SupportByVersion("Access", 9,10,11,12,14,15,16)]
	[EntityType(EntityType.IsCoClass)]
    [ComEventContract(typeof(EventContracts._OptionButtonInOptionEvents), typeof(EventContracts.DispOptionButtonEvents))]
	[TypeId("BC9E435A-F037-11CD-8701-00AA003F0F07")]
    public interface _OptionButtonInOption : _OptionButton, IEventBinding
	{
		#region Events

		/// <summary>
		/// SupportByVersion Access 9 10 11 12 14 15,16
		/// </summary>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		event _OptionButtonInOption_GotFocusEventHandler GotFocusEvent;

		/// <summary>
		/// SupportByVersion Access 9 10 11 12 14 15,16
		/// </summary>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		event _OptionButtonInOption_LostFocusEventHandler LostFocusEvent;

		/// <summary>
		/// SupportByVersion Access 9 10 11 12 14 15,16
		/// </summary>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		event _OptionButtonInOption_MouseDownEventHandler MouseDownEvent;

		/// <summary>
		/// SupportByVersion Access 9 10 11 12 14 15,16
		/// </summary>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		event _OptionButtonInOption_MouseMoveEventHandler MouseMoveEvent;

		/// <summary>
		/// SupportByVersion Access 9 10 11 12 14 15,16
		/// </summary>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		event _OptionButtonInOption_MouseUpEventHandler MouseUpEvent;

		/// <summary>
		/// SupportByVersion Access 9 10 11 12 14 15,16
		/// </summary>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		event _OptionButtonInOption_KeyDownEventHandler KeyDownEvent;

		/// <summary>
		/// SupportByVersion Access 9 10 11 12 14 15,16
		/// </summary>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		event _OptionButtonInOption_KeyPressEventHandler KeyPressEvent;

		/// <summary>
		/// SupportByVersion Access 9 10 11 12 14 15,16
		/// </summary>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		event _OptionButtonInOption_KeyUpEventHandler KeyUpEvent;

		/// <summary>
		/// SupportByVersion Access 12 14 15,16
		/// </summary>
		[SupportByVersion("Access", 12,14,15,16)]
		event _OptionButtonInOption_ClickEventHandler ClickEvent;

		/// <summary>
		/// SupportByVersion Access 12 14 15,16
		/// </summary>
		[SupportByVersion("Access", 12,14,15,16)]
		event _OptionButtonInOption_BeforeUpdateEventHandler BeforeUpdateEvent;

		/// <summary>
		/// SupportByVersion Access 12 14 15,16
		/// </summary>
		[SupportByVersion("Access", 12,14,15,16)]
		event _OptionButtonInOption_AfterUpdateEventHandler AfterUpdateEvent;

		/// <summary>
		/// SupportByVersion Access 12 14 15,16
		/// </summary>
		[SupportByVersion("Access", 12,14,15,16)]
		event _OptionButtonInOption_EnterEventHandler EnterEvent;

		/// <summary>
		/// SupportByVersion Access 12 14 15,16
		/// </summary>
		[SupportByVersion("Access", 12,14,15,16)]
		event _OptionButtonInOption_ExitEventHandler ExitEvent;

		/// <summary>
		/// SupportByVersion Access 12 14 15,16
		/// </summary>
		[SupportByVersion("Access", 12,14,15,16)]
		event _OptionButtonInOption_DblClickEventHandler DblClickEvent;

		#endregion
	}
}
