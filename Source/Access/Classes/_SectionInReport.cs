using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.AccessApi
{
	#region Delegates

	#pragma warning disable
	public delegate void _SectionInReport_FormatEventHandler(ref Int16 cancel, ref Int16 formatCount);
	public delegate void _SectionInReport_PrintEventHandler(ref Int16 cancel, ref Int16 printCount);
	public delegate void _SectionInReport_RetreatEventHandler();
	public delegate void _SectionInReport_ClickEventHandler();
	public delegate void _SectionInReport_DblClickEventHandler(ref Int16 cancel);
	public delegate void _SectionInReport_MouseDownEventHandler(ref Int16 button, ref Int16 shift, ref Single x, ref Single y);
	public delegate void _SectionInReport_MouseMoveEventHandler(ref Int16 button, ref Int16 shift, ref Single x, ref Single y);
	public delegate void _SectionInReport_MouseUpEventHandler(ref Int16 button, ref Int16 shift, ref Single x, ref Single y);
	public delegate void _SectionInReport_PaintEventHandler();
	#pragma warning restore

	#endregion

	/// <summary>
	/// CoClass _SectionInReport
	/// SupportByVersion Access, 9,10,11,12,14,15,16
	/// </summary>
	[SupportByVersion("Access", 9,10,11,12,14,15,16)]
	[EntityType(EntityType.IsCoClass)]
    [ComEventContract(typeof(EventContracts._SectionInReportEvents), typeof(EventContracts.DispSectionInReportEvents))]
	[TypeId("BC9E4360-F037-11CD-8701-00AA003F0F07")]
    public interface _SectionInReport : _Section, IEventBinding
	{
		#region Events

		/// <summary>
		/// SupportByVersion Access 9 10 11 12 14 15,16
		/// </summary>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		event _SectionInReport_FormatEventHandler FormatEvent;

		/// <summary>
		/// SupportByVersion Access 9 10 11 12 14 15,16
		/// </summary>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		event _SectionInReport_PrintEventHandler PrintEvent;

		/// <summary>
		/// SupportByVersion Access 9 10 11 12 14 15,16
		/// </summary>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		event _SectionInReport_RetreatEventHandler RetreatEvent;

		/// <summary>
		/// SupportByVersion Access 12 14 15,16
		/// </summary>
		[SupportByVersion("Access", 12,14,15,16)]
		event _SectionInReport_ClickEventHandler ClickEvent;

		/// <summary>
		/// SupportByVersion Access 12 14 15,16
		/// </summary>
		[SupportByVersion("Access", 12,14,15,16)]
		event _SectionInReport_DblClickEventHandler DblClickEvent;

		/// <summary>
		/// SupportByVersion Access 12 14 15,16
		/// </summary>
		[SupportByVersion("Access", 12,14,15,16)]
		event _SectionInReport_MouseDownEventHandler MouseDownEvent;

		/// <summary>
		/// SupportByVersion Access 12 14 15,16
		/// </summary>
		[SupportByVersion("Access", 12,14,15,16)]
		event _SectionInReport_MouseMoveEventHandler MouseMoveEvent;

		/// <summary>
		/// SupportByVersion Access 12 14 15,16
		/// </summary>
		[SupportByVersion("Access", 12,14,15,16)]
		event _SectionInReport_MouseUpEventHandler MouseUpEvent;

		/// <summary>
		/// SupportByVersion Access 12 14 15,16
		/// </summary>
		[SupportByVersion("Access", 12,14,15,16)]
		event _SectionInReport_PaintEventHandler PaintEvent;

		#endregion
	}
}
