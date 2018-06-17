using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.AccessApi
{
	#region Delegates

	#pragma warning disable
	public delegate void ReportOld_OpenEventHandler(ref Int16 cancel);
	public delegate void ReportOld_CloseEventHandler();
	public delegate void ReportOld_ActivateEventHandler();
	public delegate void ReportOld_DeactivateEventHandler();
	public delegate void ReportOld_ErrorEventHandler(ref Int16 dataErr, ref Int16 response);
	public delegate void ReportOld_NoDataEventHandler(ref Int16 cancel);
	public delegate void ReportOld_PageEventHandler();
	#pragma warning restore

	#endregion

	/// <summary>
	/// CoClass ReportOld 
	/// SupportByVersion Access, 10,11,12,14,15,16
	/// </summary>
	[SupportByVersion("Access", 10,11,12,14,15,16)]
	[EntityType(EntityType.IsCoClass)]
    [ComEventContract(typeof(EventContracts._ReportEvents))]
	[TypeId("27CE30A0-91FF-101B-AF4E-00AA003F0F07")]
    public interface ReportOld : _Report, IEventBinding
	{
		#region Events

		/// <summary>
		/// SupportByVersion Access 9 10 11 12 14 15,16
		/// </summary>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		event ReportOld_OpenEventHandler OpenEvent;

		/// <summary>
		/// SupportByVersion Access 9 10 11 12 14 15,16
		/// </summary>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		event ReportOld_CloseEventHandler CloseEvent;

		/// <summary>
		/// SupportByVersion Access 9 10 11 12 14 15,16
		/// </summary>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		event ReportOld_ActivateEventHandler ActivateEvent;

		/// <summary>
		/// SupportByVersion Access 9 10 11 12 14 15,16
		/// </summary>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		event ReportOld_DeactivateEventHandler DeactivateEvent;

		/// <summary>
		/// SupportByVersion Access 9 10 11 12 14 15,16
		/// </summary>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		event ReportOld_ErrorEventHandler ErrorEvent;

		/// <summary>
		/// SupportByVersion Access 9 10 11 12 14 15,16
		/// </summary>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		event ReportOld_NoDataEventHandler NoDataEvent;

		/// <summary>
		/// SupportByVersion Access 9 10 11 12 14 15,16
		/// </summary>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		event ReportOld_PageEventHandler PageEvent;

		#endregion
	}
}
