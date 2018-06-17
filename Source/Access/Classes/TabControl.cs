using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.AccessApi
{
	#region Delegates

	#pragma warning disable
	public delegate void TabControl_ClickEventHandler();
	public delegate void TabControl_DblClickEventHandler(ref Int16 cancel);
	public delegate void TabControl_MouseDownEventHandler(ref Int16 button, ref Int16 shift, ref Single x, ref Single y);
	public delegate void TabControl_MouseMoveEventHandler(ref Int16 button, ref Int16 shift, ref Single x, ref Single y);
	public delegate void TabControl_MouseUpEventHandler(ref Int16 button, ref Int16 shift, ref Single x, ref Single y);
	public delegate void TabControl_KeyDownEventHandler(ref Int16 keyCode, ref Int16 shift);
	public delegate void TabControl_KeyPressEventHandler(ref Int16 keyAscii);
	public delegate void TabControl_KeyUpEventHandler(ref Int16 keyCode, ref Int16 shift);
	public delegate void TabControl_ChangeEventHandler();
	#pragma warning restore

	#endregion

	/// <summary>
	/// CoClass TabControl 
	/// SupportByVersion Access, 9,10,11,12,14,15,16
	/// </summary>
	/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff844930.aspx </remarks>
	[SupportByVersion("Access", 9,10,11,12,14,15,16)]
	[EntityType(EntityType.IsCoClass)]
    [ComEventContract(typeof(EventContracts._TabControlEvents), typeof(EventContracts.DispTabControlEvents))]
	[TypeId("3B06E970-E47C-11CD-8701-00AA003F0F07")]
    public interface TabControl : _TabControl, IEventBinding
	{
		#region Events

		/// <summary>
		/// SupportByVersion Access 9 10 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff191701.aspx </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		event TabControl_ClickEventHandler ClickEvent;

		/// <summary>
		/// SupportByVersion Access 9 10 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff834998.aspx </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		event TabControl_DblClickEventHandler DblClickEvent;

		/// <summary>
		/// SupportByVersion Access 9 10 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff196779.aspx </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		event TabControl_MouseDownEventHandler MouseDownEvent;

		/// <summary>
		/// SupportByVersion Access 9 10 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff835969.aspx </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		event TabControl_MouseMoveEventHandler MouseMoveEvent;

		/// <summary>
		/// SupportByVersion Access 9 10 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff192315.aspx </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		event TabControl_MouseUpEventHandler MouseUpEvent;

		/// <summary>
		/// SupportByVersion Access 9 10 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff844721.aspx </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		event TabControl_KeyDownEventHandler KeyDownEvent;

		/// <summary>
		/// SupportByVersion Access 9 10 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff193508.aspx </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		event TabControl_KeyPressEventHandler KeyPressEvent;

		/// <summary>
		/// SupportByVersion Access 9 10 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff823051.aspx </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		event TabControl_KeyUpEventHandler KeyUpEvent;

		/// <summary>
		/// SupportByVersion Access 9 10 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff835955.aspx </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		event TabControl_ChangeEventHandler ChangeEvent;

		#endregion
	}
}
