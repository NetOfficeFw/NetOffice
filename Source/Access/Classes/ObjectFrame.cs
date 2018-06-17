using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.AccessApi
{
	#region Delegates

	#pragma warning disable
	public delegate void ObjectFrame_UpdatedEventHandler(ref Int16 code);
	public delegate void ObjectFrame_EnterEventHandler();
	public delegate void ObjectFrame_ExitEventHandler(ref Int16 cancel);
	public delegate void ObjectFrame_GotFocusEventHandler();
	public delegate void ObjectFrame_LostFocusEventHandler();
	public delegate void ObjectFrame_ClickEventHandler();
	public delegate void ObjectFrame_DblClickEventHandler(ref Int16 cancel);
	public delegate void ObjectFrame_MouseDownEventHandler(ref Int16 button, ref Int16 shift, ref Single x, ref Single y);
	public delegate void ObjectFrame_MouseMoveEventHandler(ref Int16 button, ref Int16 shift, ref Single x, ref Single y);
	public delegate void ObjectFrame_MouseUpEventHandler(ref Int16 button, ref Int16 shift, ref Single x, ref Single y);
	#pragma warning restore

	#endregion

	/// <summary>
	/// CoClass ObjectFrame 
	/// SupportByVersion Access, 9,10,11,12,14,15,16
	/// </summary>
	/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff845258.aspx </remarks>
	[SupportByVersion("Access", 9,10,11,12,14,15,16)]
	[EntityType(EntityType.IsCoClass)]
    [ComEventContract(typeof(EventContracts._ObjectFrameEvents), typeof(EventContracts.DispObjectFrameEvents))]
	[TypeId("3B06E95D-E47C-11CD-8701-00AA003F0F07")]
    public interface ObjectFrame : _ObjectFrame, IEventBinding
	{
		#region Events

		/// <summary>
		/// SupportByVersion Access 9 10 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff196691.aspx </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		event ObjectFrame_UpdatedEventHandler UpdatedEvent;

		/// <summary>
		/// SupportByVersion Access 9 10 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff196796.aspx </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		event ObjectFrame_EnterEventHandler EnterEvent;

		/// <summary>
		/// SupportByVersion Access 9 10 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff197980.aspx </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		event ObjectFrame_ExitEventHandler ExitEvent;

		/// <summary>
		/// SupportByVersion Access 9 10 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff192894.aspx </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		event ObjectFrame_GotFocusEventHandler GotFocusEvent;

		/// <summary>
		/// SupportByVersion Access 9 10 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff835022.aspx </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		event ObjectFrame_LostFocusEventHandler LostFocusEvent;

		/// <summary>
		/// SupportByVersion Access 9 10 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff196135.aspx </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		event ObjectFrame_ClickEventHandler ClickEvent;

		/// <summary>
		/// SupportByVersion Access 9 10 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff196751.aspx </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		event ObjectFrame_DblClickEventHandler DblClickEvent;

		/// <summary>
		/// SupportByVersion Access 9 10 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff194145.aspx </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		event ObjectFrame_MouseDownEventHandler MouseDownEvent;

		/// <summary>
		/// SupportByVersion Access 9 10 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff821466.aspx </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		event ObjectFrame_MouseMoveEventHandler MouseMoveEvent;

		/// <summary>
		/// SupportByVersion Access 9 10 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff845282.aspx </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		event ObjectFrame_MouseUpEventHandler MouseUpEvent;

		#endregion
	}
}
