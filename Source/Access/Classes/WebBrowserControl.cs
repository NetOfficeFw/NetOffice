using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.AccessApi
{
	#region Delegates

	#pragma warning disable
	public delegate void WebBrowserControl_UpdatedEventHandler(ref Int16 code);
	public delegate void WebBrowserControl_BeforeUpdateEventHandler(ref Int16 cancel);
	public delegate void WebBrowserControl_AfterUpdateEventHandler();
	public delegate void WebBrowserControl_EnterEventHandler();
	public delegate void WebBrowserControl_ExitEventHandler(ref Int16 cancel);
	public delegate void WebBrowserControl_DirtyEventHandler(ref Int16 cancel);
	public delegate void WebBrowserControl_ChangeEventHandler();
	public delegate void WebBrowserControl_GotFocusEventHandler();
	public delegate void WebBrowserControl_LostFocusEventHandler();
	public delegate void WebBrowserControl_ClickEventHandler();
	public delegate void WebBrowserControl_DblClickEventHandler(ref Int16 cancel);
	public delegate void WebBrowserControl_MouseDownEventHandler(ref Int16 button, ref Int16 shift, ref Single x, ref Single y);
	public delegate void WebBrowserControl_MouseMoveEventHandler(ref Int16 button, ref Int16 shift, ref Single x, ref Single y);
	public delegate void WebBrowserControl_MouseUpEventHandler(ref Int16 button, ref Int16 shift, ref Single x, ref Single y);
	public delegate void WebBrowserControl_KeyDownEventHandler(ref Int16 keyCode, ref Int16 shift);
	public delegate void WebBrowserControl_KeyPressEventHandler(ref Int16 keyAscii);
	public delegate void WebBrowserControl_KeyUpEventHandler(ref Int16 keyCode, ref Int16 shift);
	public delegate void WebBrowserControl_BeforeNavigate2EventHandler(ICOMObject pDisp, ref object url, ref object flags, ref object targetFrameName, ref object postData, ref object headers, ref bool cancel);
	public delegate void WebBrowserControl_DocumentCompleteEventHandler(ICOMObject pDisp, ref object url);
	public delegate void WebBrowserControl_ProgressChangeEventHandler(Int32 progress, Int32 progressMax);
	public delegate void WebBrowserControl_NavigateErrorEventHandler(ICOMObject pDisp, ref object url, ref object targetFrameName, ref object satusCode, ref bool cancel);
	#pragma warning restore

	#endregion

	/// <summary>
	/// CoClass WebBrowserControl 
	/// SupportByVersion Access, 14,15,16
	/// </summary>
	/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff835067.aspx </remarks>
	[SupportByVersion("Access", 14,15,16)]
	[EntityType(EntityType.IsCoClass)]
    [ComEventContract(typeof(EventContracts.DispWebBrowserControlEvents))]
	[TypeId("D303AC37-74DB-45B9-8C22-AD7C3FBA68EF")]
    public interface WebBrowserControl : _WebBrowserControl, IEventBinding
	{
		#region Events

		/// <summary>
		/// SupportByVersion Access 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff196764.aspx </remarks>
		[SupportByVersion("Access", 14,15,16)]
		event WebBrowserControl_UpdatedEventHandler UpdatedEvent;

		/// <summary>
		/// SupportByVersion Access 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff195884.aspx </remarks>
		[SupportByVersion("Access", 14,15,16)]
		event WebBrowserControl_BeforeUpdateEventHandler BeforeUpdateEvent;

		/// <summary>
		/// SupportByVersion Access 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff197400.aspx </remarks>
		[SupportByVersion("Access", 14,15,16)]
		event WebBrowserControl_AfterUpdateEventHandler AfterUpdateEvent;

		/// <summary>
		/// SupportByVersion Access 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff193153.aspx </remarks>
		[SupportByVersion("Access", 14,15,16)]
		event WebBrowserControl_EnterEventHandler EnterEvent;

		/// <summary>
		/// SupportByVersion Access 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff821106.aspx </remarks>
		[SupportByVersion("Access", 14,15,16)]
		event WebBrowserControl_ExitEventHandler ExitEvent;

		/// <summary>
		/// SupportByVersion Access 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff192440.aspx </remarks>
		[SupportByVersion("Access", 14,15,16)]
		event WebBrowserControl_DirtyEventHandler DirtyEvent;

		/// <summary>
		/// SupportByVersion Access 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff192510.aspx </remarks>
		[SupportByVersion("Access", 14,15,16)]
		event WebBrowserControl_ChangeEventHandler ChangeEvent;

		/// <summary>
		/// SupportByVersion Access 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff195783.aspx </remarks>
		[SupportByVersion("Access", 14,15,16)]
		event WebBrowserControl_GotFocusEventHandler GotFocusEvent;

		/// <summary>
		/// SupportByVersion Access 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff193588.aspx </remarks>
		[SupportByVersion("Access", 14,15,16)]
		event WebBrowserControl_LostFocusEventHandler LostFocusEvent;

		/// <summary>
		/// SupportByVersion Access 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff192861.aspx </remarks>
		[SupportByVersion("Access", 14,15,16)]
		event WebBrowserControl_ClickEventHandler ClickEvent;

		/// <summary>
		/// SupportByVersion Access 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff835690.aspx </remarks>
		[SupportByVersion("Access", 14,15,16)]
		event WebBrowserControl_DblClickEventHandler DblClickEvent;

		/// <summary>
		/// SupportByVersion Access 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff823017.aspx </remarks>
		[SupportByVersion("Access", 14,15,16)]
		event WebBrowserControl_MouseDownEventHandler MouseDownEvent;

		/// <summary>
		/// SupportByVersion Access 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff845665.aspx </remarks>
		[SupportByVersion("Access", 14,15,16)]
		event WebBrowserControl_MouseMoveEventHandler MouseMoveEvent;

		/// <summary>
		/// SupportByVersion Access 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff196763.aspx </remarks>
		[SupportByVersion("Access", 14,15,16)]
		event WebBrowserControl_MouseUpEventHandler MouseUpEvent;

		/// <summary>
		/// SupportByVersion Access 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff845359.aspx </remarks>
		[SupportByVersion("Access", 14,15,16)]
		event WebBrowserControl_KeyDownEventHandler KeyDownEvent;

		/// <summary>
		/// SupportByVersion Access 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff194971.aspx </remarks>
		[SupportByVersion("Access", 14,15,16)]
		event WebBrowserControl_KeyPressEventHandler KeyPressEvent;

		/// <summary>
		/// SupportByVersion Access 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff835380.aspx </remarks>
		[SupportByVersion("Access", 14,15,16)]
		event WebBrowserControl_KeyUpEventHandler KeyUpEvent;

		/// <summary>
		/// SupportByVersion Access 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff196461.aspx </remarks>
		[SupportByVersion("Access", 14,15,16)]
		event WebBrowserControl_BeforeNavigate2EventHandler BeforeNavigate2Event;

		/// <summary>
		/// SupportByVersion Access 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff197343.aspx </remarks>
		[SupportByVersion("Access", 14,15,16)]
		event WebBrowserControl_DocumentCompleteEventHandler DocumentCompleteEvent;

		/// <summary>
		/// SupportByVersion Access 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff845660.aspx </remarks>
		[SupportByVersion("Access", 14,15,16)]
		event WebBrowserControl_ProgressChangeEventHandler ProgressChangeEvent;

		/// <summary>
		/// SupportByVersion Access 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff845715.aspx </remarks>
		[SupportByVersion("Access", 14,15,16)]
		event WebBrowserControl_NavigateErrorEventHandler NavigateErrorEvent;

		#endregion
	}
}
