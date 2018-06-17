using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.AccessApi
{
	#region Delegates

	#pragma warning disable
	public delegate void NavigationControl_BeforeUpdateEventHandler(ref Int16 cancel);
	public delegate void NavigationControl_AfterUpdateEventHandler();
	public delegate void NavigationControl_ChangeEventHandler();
	public delegate void NavigationControl_EnterEventHandler();
	public delegate void NavigationControl_ExitEventHandler(ref Int16 cancel);
	public delegate void NavigationControl_GotFocusEventHandler();
	public delegate void NavigationControl_LostFocusEventHandler();
	public delegate void NavigationControl_ClickEventHandler();
	public delegate void NavigationControl_DblClickEventHandler(ref Int16 cancel);
	public delegate void NavigationControl_MouseDownEventHandler(ref Int16 button, ref Int16 shift, ref Single x, ref Single y);
	public delegate void NavigationControl_MouseMoveEventHandler(ref Int16 button, ref Int16 shift, ref Single x, ref Single y);
	public delegate void NavigationControl_MouseUpEventHandler(ref Int16 button, ref Int16 shift, ref Single x, ref Single y);
	public delegate void NavigationControl_KeyDownEventHandler(ref Int16 keyCode, ref Int16 shift);
	public delegate void NavigationControl_KeyPressEventHandler(ref Int16 keyAscii);
	public delegate void NavigationControl_KeyUpEventHandler(ref Int16 keyCode, ref Int16 shift);
	public delegate void NavigationControl_DirtyEventHandler(ref Int16 cancel);
	public delegate void NavigationControl_UndoEventHandler(ref Int16 cancel);
	#pragma warning restore

	#endregion

	/// <summary>
	/// CoClass NavigationControl 
	/// SupportByVersion Access, 14,15,16
	/// </summary>
	/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff821468.aspx </remarks>
	[SupportByVersion("Access", 14,15,16)]
	[EntityType(EntityType.IsCoClass)]
    [ComEventContract(typeof(EventContracts.DispNavigationControlEvents))]
	[TypeId("3B06E989-E47C-11CD-8701-00AA003F0F07")]
    public interface NavigationControl : _NavigationControl, IEventBinding
	{
		#region Events

		/// <summary>
		/// SupportByVersion Access 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff192533.aspx </remarks>
		[SupportByVersion("Access", 14,15,16)]
		event NavigationControl_BeforeUpdateEventHandler BeforeUpdateEvent;

		/// <summary>
		/// SupportByVersion Access 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff821744.aspx </remarks>
		[SupportByVersion("Access", 14,15,16)]
		event NavigationControl_AfterUpdateEventHandler AfterUpdateEvent;

		/// <summary>
		/// SupportByVersion Access 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff192951.aspx </remarks>
		[SupportByVersion("Access", 14,15,16)]
		event NavigationControl_ChangeEventHandler ChangeEvent;

		/// <summary>
		/// SupportByVersion Access 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff192267.aspx </remarks>
		[SupportByVersion("Access", 14,15,16)]
		event NavigationControl_EnterEventHandler EnterEvent;

		/// <summary>
		/// SupportByVersion Access 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff193801.aspx </remarks>
		[SupportByVersion("Access", 14,15,16)]
		event NavigationControl_ExitEventHandler ExitEvent;

		/// <summary>
		/// SupportByVersion Access 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff193831.aspx </remarks>
		[SupportByVersion("Access", 14,15,16)]
		event NavigationControl_GotFocusEventHandler GotFocusEvent;

		/// <summary>
		/// SupportByVersion Access 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff194817.aspx </remarks>
		[SupportByVersion("Access", 14,15,16)]
		event NavigationControl_LostFocusEventHandler LostFocusEvent;

		/// <summary>
		/// SupportByVersion Access 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff823080.aspx </remarks>
		[SupportByVersion("Access", 14,15,16)]
		event NavigationControl_ClickEventHandler ClickEvent;

		/// <summary>
		/// SupportByVersion Access 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff836976.aspx </remarks>
		[SupportByVersion("Access", 14,15,16)]
		event NavigationControl_DblClickEventHandler DblClickEvent;

		/// <summary>
		/// SupportByVersion Access 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff844823.aspx </remarks>
		[SupportByVersion("Access", 14,15,16)]
		event NavigationControl_MouseDownEventHandler MouseDownEvent;

		/// <summary>
		/// SupportByVersion Access 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff821140.aspx </remarks>
		[SupportByVersion("Access", 14,15,16)]
		event NavigationControl_MouseMoveEventHandler MouseMoveEvent;

		/// <summary>
		/// SupportByVersion Access 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff845572.aspx </remarks>
		[SupportByVersion("Access", 14,15,16)]
		event NavigationControl_MouseUpEventHandler MouseUpEvent;

		/// <summary>
		/// SupportByVersion Access 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff844778.aspx </remarks>
		[SupportByVersion("Access", 14,15,16)]
		event NavigationControl_KeyDownEventHandler KeyDownEvent;

		/// <summary>
		/// SupportByVersion Access 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff835989.aspx </remarks>
		[SupportByVersion("Access", 14,15,16)]
		event NavigationControl_KeyPressEventHandler KeyPressEvent;

		/// <summary>
		/// SupportByVersion Access 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff192466.aspx </remarks>
		[SupportByVersion("Access", 14,15,16)]
		event NavigationControl_KeyUpEventHandler KeyUpEvent;

		/// <summary>
		/// SupportByVersion Access 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff194860.aspx </remarks>
		[SupportByVersion("Access", 14,15,16)]
		event NavigationControl_DirtyEventHandler DirtyEvent;

		/// <summary>
		/// SupportByVersion Access 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff836273.aspx </remarks>
		[SupportByVersion("Access", 14,15,16)]
		event NavigationControl_UndoEventHandler UndoEvent;

		#endregion
	}
}
