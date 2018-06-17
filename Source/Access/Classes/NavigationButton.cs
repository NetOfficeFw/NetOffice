using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.AccessApi
{
	#region Delegates

	#pragma warning disable
	public delegate void NavigationButton_ClickEventHandler();
	public delegate void NavigationButton_EnterEventHandler();
	public delegate void NavigationButton_ExitEventHandler(ref Int16 cancel);
	public delegate void NavigationButton_GotFocusEventHandler();
	public delegate void NavigationButton_LostFocusEventHandler();
	public delegate void NavigationButton_DblClickEventHandler(ref Int16 cancel);
	public delegate void NavigationButton_MouseDownEventHandler(ref Int16 button, ref Int16 shift, ref Single x, ref Single y);
	public delegate void NavigationButton_MouseMoveEventHandler(ref Int16 button, ref Int16 shift, ref Single x, ref Single y);
	public delegate void NavigationButton_MouseUpEventHandler(ref Int16 button, ref Int16 shift, ref Single x, ref Single y);
	public delegate void NavigationButton_KeyDownEventHandler(ref Int16 keyCode, ref Int16 shift);
	public delegate void NavigationButton_KeyPressEventHandler(ref Int16 keyAscii);
	public delegate void NavigationButton_KeyUpEventHandler(ref Int16 keyCode, ref Int16 shift);
	#pragma warning restore

	#endregion

	/// <summary>
	/// CoClass NavigationButton 
	/// SupportByVersion Access, 14,15,16
	/// </summary>
	/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff821707.aspx </remarks>
	[SupportByVersion("Access", 14,15,16)]
	[EntityType(EntityType.IsCoClass)]
    [ComEventContract(typeof(EventContracts.DispNavigationButtonEvents))]
	[TypeId("3B06E993-E47C-11CD-8701-00AA003F0F07")]
    public interface NavigationButton : _NavigationButton, IEventBinding
	{
		#region Events

		/// <summary>
		/// SupportByVersion Access 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff822048.aspx </remarks>
		[SupportByVersion("Access", 14,15,16)]
		event NavigationButton_ClickEventHandler ClickEvent;

		/// <summary>
		/// SupportByVersion Access 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff822726.aspx </remarks>
		[SupportByVersion("Access", 14,15,16)]
		event NavigationButton_EnterEventHandler EnterEvent;

		/// <summary>
		/// SupportByVersion Access 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff196059.aspx </remarks>
		[SupportByVersion("Access", 14,15,16)]
		event NavigationButton_ExitEventHandler ExitEvent;

		/// <summary>
		/// SupportByVersion Access 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff192653.aspx </remarks>
		[SupportByVersion("Access", 14,15,16)]
		event NavigationButton_GotFocusEventHandler GotFocusEvent;

		/// <summary>
		/// SupportByVersion Access 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff835383.aspx </remarks>
		[SupportByVersion("Access", 14,15,16)]
		event NavigationButton_LostFocusEventHandler LostFocusEvent;

		/// <summary>
		/// SupportByVersion Access 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff820826.aspx </remarks>
		[SupportByVersion("Access", 14,15,16)]
		event NavigationButton_DblClickEventHandler DblClickEvent;

		/// <summary>
		/// SupportByVersion Access 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff197983.aspx </remarks>
		[SupportByVersion("Access", 14,15,16)]
		event NavigationButton_MouseDownEventHandler MouseDownEvent;

		/// <summary>
		/// SupportByVersion Access 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff195891.aspx </remarks>
		[SupportByVersion("Access", 14,15,16)]
		event NavigationButton_MouseMoveEventHandler MouseMoveEvent;

		/// <summary>
		/// SupportByVersion Access 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff845798.aspx </remarks>
		[SupportByVersion("Access", 14,15,16)]
		event NavigationButton_MouseUpEventHandler MouseUpEvent;

		/// <summary>
		/// SupportByVersion Access 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff837321.aspx </remarks>
		[SupportByVersion("Access", 14,15,16)]
		event NavigationButton_KeyDownEventHandler KeyDownEvent;

		/// <summary>
		/// SupportByVersion Access 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff197089.aspx </remarks>
		[SupportByVersion("Access", 14,15,16)]
		event NavigationButton_KeyPressEventHandler KeyPressEvent;

		/// <summary>
		/// SupportByVersion Access 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff195130.aspx </remarks>
		[SupportByVersion("Access", 14,15,16)]
		event NavigationButton_KeyUpEventHandler KeyUpEvent;

		#endregion
	}
}
