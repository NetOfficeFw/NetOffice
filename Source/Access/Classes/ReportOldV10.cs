using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.AccessApi
{
	#region Delegates

	#pragma warning disable
	public delegate void ReportOldV10_OpenEventHandler(ref Int16 cancel);
	public delegate void ReportOldV10_CloseEventHandler();
	public delegate void ReportOldV10_ActivateEventHandler();
	public delegate void ReportOldV10_DeactivateEventHandler();
	public delegate void ReportOldV10_ErrorEventHandler(ref Int16 dataErr, ref Int16 response);
	public delegate void ReportOldV10_NoDataEventHandler(ref Int16 cancel);
	public delegate void ReportOldV10_PageEventHandler();
	#pragma warning restore

	#endregion

	/// <summary>
	/// CoClass ReportOldV10 
	/// SupportByVersion Access, 12,14,15,16
	/// </summary>
	[SupportByVersion("Access", 12,14,15,16)]
	[EntityType(EntityType.IsCoClass)]
    [ComEventContract(typeof(EventContracts._ReportEvents))]
	[TypeId("ECD1EADA-D373-11D3-8D21-0050048383FB")]
    public interface ReportOldV10 : _Report2, IEventBinding
	{
		#region Events

		/// <summary>
		/// SupportByVersion Access 9 10 11 12 14 15,16
		/// </summary>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		event ReportOldV10_OpenEventHandler OpenEvent;

		/// <summary>
		/// SupportByVersion Access 9 10 11 12 14 15,16
		/// </summary>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		event ReportOldV10_CloseEventHandler CloseEvent;

		/// <summary>
		/// SupportByVersion Access 9 10 11 12 14 15,16
		/// </summary>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		event ReportOldV10_ActivateEventHandler ActivateEvent;

		/// <summary>
		/// SupportByVersion Access 9 10 11 12 14 15,16
		/// </summary>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		event ReportOldV10_DeactivateEventHandler DeactivateEvent;

		/// <summary>
		/// SupportByVersion Access 9 10 11 12 14 15,16
		/// </summary>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		event ReportOldV10_ErrorEventHandler ErrorEvent;

		/// <summary>
		/// SupportByVersion Access 9 10 11 12 14 15,16
		/// </summary>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		event ReportOldV10_NoDataEventHandler NoDataEvent;

		/// <summary>
		/// SupportByVersion Access 9 10 11 12 14 15,16
		/// </summary>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		event ReportOldV10_PageEventHandler PageEvent;

		#endregion
	}
}
