using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.AccessApi
{
	#region Delegates

	#pragma warning disable
	public delegate void Page_ClickEventHandler();
	public delegate void Page_DblClickEventHandler(ref Int16 cancel);
	public delegate void Page_MouseDownEventHandler(ref Int16 button, ref Int16 shift, ref Single x, ref Single y);
	public delegate void Page_MouseMoveEventHandler(ref Int16 button, ref Int16 shift, ref Single x, ref Single y);
	public delegate void Page_MouseUpEventHandler(ref Int16 button, ref Int16 shift, ref Single x, ref Single y);
	#pragma warning restore

	#endregion

	/// <summary>
	/// CoClass Page 
	/// SupportByVersion Access, 9,10,11,12,14,15,16
	/// </summary>
	/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194955.aspx </remarks>
	[SupportByVersion("Access", 9,10,11,12,14,15,16)]
	[EntityType(EntityType.IsCoClass)]
    [ComEventContract(typeof(EventContracts._PageEvents), typeof(EventContracts.DispPageEvents))]
	[TypeId("3B06E973-E47C-11CD-8701-00AA003F0F07")]
    public interface Page : _Page, IEventBinding
	{
		#region Events

		/// <summary>
		/// SupportByVersion Access 9 10 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff196433.aspx </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		event Page_ClickEventHandler ClickEvent;

		/// <summary>
		/// SupportByVersion Access 9 10 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff198356.aspx </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		event Page_DblClickEventHandler DblClickEvent;

		/// <summary>
		/// SupportByVersion Access 9 10 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff195462.aspx </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		event Page_MouseDownEventHandler MouseDownEvent;

		/// <summary>
		/// SupportByVersion Access 9 10 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff820745.aspx </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		event Page_MouseMoveEventHandler MouseMoveEvent;

		/// <summary>
		/// SupportByVersion Access 9 10 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff835950.aspx </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		event Page_MouseUpEventHandler MouseUpEvent;

		#endregion
	}
}
