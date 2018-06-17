using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.AccessApi
{
	/// <summary>
	/// CoClass _CustomControlInReport
	/// SupportByVersion Access, 9,10,11,12,14,15,16
	/// </summary>
	[SupportByVersion("Access", 9,10,11,12,14,15,16)]
	[EntityType(EntityType.IsCoClass)]
    [ComEventContract(typeof(EventContracts._CustomControlInReportEvents), typeof(EventContracts.DispCustomControlInReportEvents))]
	[TypeId("300471E0-7426-11CE-AB63-00AA0042B7CE")]
    public interface _CustomControlInReport : _CustomControl
	{

	}
}
