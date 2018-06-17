using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.AccessApi
{
	#region Delegates

	#pragma warning disable
	public delegate void Report_OpenEventHandler(ref Int16 cancel);
	public delegate void Report_CloseEventHandler();
	public delegate void Report_ActivateEventHandler();
	public delegate void Report_DeactivateEventHandler();
	public delegate void Report_ErrorEventHandler(ref Int16 dataErr, ref Int16 response);
	public delegate void Report_NoDataEventHandler(ref Int16 cancel);
	public delegate void Report_PageEventHandler();
	public delegate void Report_CurrentEventHandler();
	public delegate void Report_LoadEventHandler();
	public delegate void Report_ResizeEventHandler();
	public delegate void Report_UnloadEventHandler(ref Int16 cancel);
	public delegate void Report_GotFocusEventHandler();
	public delegate void Report_LostFocusEventHandler();
	public delegate void Report_ClickEventHandler();
	public delegate void Report_DblClickEventHandler(ref Int16 cancel);
	public delegate void Report_MouseDownEventHandler(ref Int16 button, ref Int16 shift, ref Single x, ref Single y);
	public delegate void Report_MouseMoveEventHandler(ref Int16 button, ref Int16 shift, ref Single x, ref Single y);
	public delegate void Report_MouseUpEventHandler(ref Int16 button, ref Int16 shift, ref Single x, ref Single y);
	public delegate void Report_KeyDownEventHandler(ref Int16 keyCode, ref Int16 shift);
	public delegate void Report_KeyPressEventHandler(ref Int16 keyAscii);
	public delegate void Report_KeyUpEventHandler(ref Int16 keyCode, ref Int16 Shift);
	public delegate void Report_TimerEventHandler();
	public delegate void Report_FilterEventHandler(ref Int16 cancel, ref Int16 filterType);
	public delegate void Report_ApplyFilterEventHandler(ref Int16 cancel, ref Int16 applyType);
	public delegate void Report_MouseWheelEventHandler(bool page, Int32 count);
	#pragma warning restore

	#endregion

	/// <summary>
	/// CoClass Report 
	/// SupportByVersion Access, 9,10,11,12,14,15,16
	/// </summary>
	/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff195583.aspx </remarks>
	[SupportByVersion("Access", 9,10,11,12,14,15,16)]
	[EntityType(EntityType.IsCoClass)]
    [ComEventContract(typeof(EventContracts._ReportEvents), typeof(EventContracts._ReportEvents2))]
	[TypeId("27CE30A0-91FF-101B-AF4E-00AA003F0F07")]
    public interface Report : _Report3, IEventBinding
	{
		#region Events

		/// <summary>
		/// SupportByVersion Access 9 10 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff834749.aspx </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		event Report_OpenEventHandler OpenEvent;

		/// <summary>
		/// SupportByVersion Access 9 10 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff193942.aspx </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		event Report_CloseEventHandler CloseEvent;

		/// <summary>
		/// SupportByVersion Access 9 10 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff194215.aspx </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		event Report_ActivateEventHandler ActivateEvent;

		/// <summary>
		/// SupportByVersion Access 9 10 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff845512.aspx </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		event Report_DeactivateEventHandler DeactivateEvent;

		/// <summary>
		/// SupportByVersion Access 9 10 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff844940.aspx </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		event Report_ErrorEventHandler ErrorEvent;

		/// <summary>
		/// SupportByVersion Access 9 10 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff837041.aspx </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		event Report_NoDataEventHandler NoDataEvent;

		/// <summary>
		/// SupportByVersion Access 9 10 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff823057.aspx </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		event Report_PageEventHandler PageEvent;

		/// <summary>
		/// SupportByVersion Access 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff821736.aspx </remarks>
		[SupportByVersion("Access", 12,14,15,16)]
		event Report_CurrentEventHandler CurrentEvent;

		/// <summary>
		/// SupportByVersion Access 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff197739.aspx </remarks>
		[SupportByVersion("Access", 12,14,15,16)]
		event Report_LoadEventHandler LoadEvent;

		/// <summary>
		/// SupportByVersion Access 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff834460.aspx </remarks>
		[SupportByVersion("Access", 12,14,15,16)]
		event Report_ResizeEventHandler ResizeEvent;

		/// <summary>
		/// SupportByVersion Access 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff844928.aspx </remarks>
		[SupportByVersion("Access", 12,14,15,16)]
		event Report_UnloadEventHandler UnloadEvent;

		/// <summary>
		/// SupportByVersion Access 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff195218.aspx </remarks>
		[SupportByVersion("Access", 12,14,15,16)]
		event Report_GotFocusEventHandler GotFocusEvent;

		/// <summary>
		/// SupportByVersion Access 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff197321.aspx </remarks>
		[SupportByVersion("Access", 12,14,15,16)]
		event Report_LostFocusEventHandler LostFocusEvent;

		/// <summary>
		/// SupportByVersion Access 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff192496.aspx </remarks>
		[SupportByVersion("Access", 12,14,15,16)]
		event Report_ClickEventHandler ClickEvent;

		/// <summary>
		/// SupportByVersion Access 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff835945.aspx </remarks>
		[SupportByVersion("Access", 12,14,15,16)]
		event Report_DblClickEventHandler DblClickEvent;

		/// <summary>
		/// SupportByVersion Access 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff837216.aspx </remarks>
		[SupportByVersion("Access", 12,14,15,16)]
		event Report_MouseDownEventHandler MouseDownEvent;

		/// <summary>
		/// SupportByVersion Access 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff822431.aspx </remarks>
		[SupportByVersion("Access", 12,14,15,16)]
		event Report_MouseMoveEventHandler MouseMoveEvent;

		/// <summary>
		/// SupportByVersion Access 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff836025.aspx </remarks>
		[SupportByVersion("Access", 12,14,15,16)]
		event Report_MouseUpEventHandler MouseUpEvent;

		/// <summary>
		/// SupportByVersion Access 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff822041.aspx </remarks>
		[SupportByVersion("Access", 12,14,15,16)]
		event Report_KeyDownEventHandler KeyDownEvent;

		/// <summary>
		/// SupportByVersion Access 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff845166.aspx </remarks>
		[SupportByVersion("Access", 12,14,15,16)]
		event Report_KeyPressEventHandler KeyPressEvent;

		/// <summary>
		/// SupportByVersion Access 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff194162.aspx </remarks>
		[SupportByVersion("Access", 12,14,15,16)]
		event Report_KeyUpEventHandler KeyUpEvent;

		/// <summary>
		/// SupportByVersion Access 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff193962.aspx </remarks>
		[SupportByVersion("Access", 12,14,15,16)]
		event Report_TimerEventHandler TimerEvent;

		/// <summary>
		/// SupportByVersion Access 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff845429.aspx </remarks>
		[SupportByVersion("Access", 12,14,15,16)]
		event Report_FilterEventHandler FilterEvent;

		/// <summary>
		/// SupportByVersion Access 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff193193.aspx </remarks>
		[SupportByVersion("Access", 12,14,15,16)]
		event Report_ApplyFilterEventHandler ApplyFilterEvent;

		/// <summary>
		/// SupportByVersion Access 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff198093.aspx </remarks>
		[SupportByVersion("Access", 12,14,15,16)]
		event Report_MouseWheelEventHandler MouseWheelEvent;

		#endregion
	}
}
