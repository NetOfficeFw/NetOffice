using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.AccessApi
{
	#region Delegates

	#pragma warning disable
	public delegate void _PageHdrFtrInReport_FormatEventHandler(ref Int16 cancel, ref Int16 formatCount);
	public delegate void _PageHdrFtrInReport_PrintEventHandler(ref Int16 cancel, ref Int16 printCount);
	public delegate void _PageHdrFtrInReport_ClickEventHandler();
	public delegate void _PageHdrFtrInReport_DblClickEventHandler(ref Int16 cancel);
	public delegate void _PageHdrFtrInReport_MouseDownEventHandler(ref Int16 button, ref Int16 shift, ref Single x, ref Single y);
	public delegate void _PageHdrFtrInReport_MouseMoveEventHandler(ref Int16 button, ref Int16 shift, ref Single x, ref Single y);
	public delegate void _PageHdrFtrInReport_MouseUpEventHandler(ref Int16 button, ref Int16 shift, ref Single x, ref Single y);
	public delegate void _PageHdrFtrInReport_PaintEventHandler();
	#pragma warning restore

	#endregion

	/// <summary>
	/// CoClass _PageHdrFtrInReport
	/// SupportByVersion Access, 9,10,11,12,14,15,16
	/// </summary>
	[SupportByVersion("Access", 9,10,11,12,14,15,16)]
	[EntityType(EntityType.IsCoClass)]
    [ComEventContract(typeof(EventContracts._PageHdrFtrInReportEvents), typeof(EventContracts.DispPageHdrFtrInReportEvents))]
	[TypeId("7AD9E906-BAF8-11CE-A68A-00AA003F0F07")]
    public interface _PageHdrFtrInReport : _Section, IEventBinding
	{
		#region Events

		/// <summary>
		/// SupportByVersion Access 9 10 11 12 14 15,16
		/// </summary>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		event _PageHdrFtrInReport_FormatEventHandler FormatEvent;

		/// <summary>
		/// SupportByVersion Access 9 10 11 12 14 15,16
		/// </summary>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		event _PageHdrFtrInReport_PrintEventHandler PrintEvent;

		/// <summary>
		/// SupportByVersion Access 12 14 15,16
		/// </summary>
		[SupportByVersion("Access", 12,14,15,16)]
		event _PageHdrFtrInReport_ClickEventHandler ClickEvent;

		/// <summary>
		/// SupportByVersion Access 12 14 15,16
		/// </summary>
		[SupportByVersion("Access", 12,14,15,16)]
		event _PageHdrFtrInReport_DblClickEventHandler DblClickEvent;

		/// <summary>
		/// SupportByVersion Access 12 14 15,16
		/// </summary>
		[SupportByVersion("Access", 12,14,15,16)]
		event _PageHdrFtrInReport_MouseDownEventHandler MouseDownEvent;

		/// <summary>
		/// SupportByVersion Access 12 14 15,16
		/// </summary>
		[SupportByVersion("Access", 12,14,15,16)]
		event _PageHdrFtrInReport_MouseMoveEventHandler MouseMoveEvent;

		/// <summary>
		/// SupportByVersion Access 12 14 15,16
		/// </summary>
		[SupportByVersion("Access", 12,14,15,16)]
		event _PageHdrFtrInReport_MouseUpEventHandler MouseUpEvent;

		/// <summary>
		/// SupportByVersion Access 12 14 15,16
		/// </summary>
		[SupportByVersion("Access", 12,14,15,16)]
		event _PageHdrFtrInReport_PaintEventHandler PaintEvent;

		#endregion
	}
}
