using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.AccessApi
{
	/// <summary>
	/// CoClass _ControlInReportEvents
	/// SupportByVersion Access, 9,10,11,12,14,15,16
	/// </summary>
	[SupportByVersion("Access", 9,10,11,12,14,15,16)]
	[EntityType(EntityType.IsCoClass)]
    [ComEventContract(typeof(EventContracts.__ControlInReportEvents), typeof(EventContracts._DispControlInReportEvents))]
	[TypeId("90B322A4-F1D9-11CD-8701-00AA003F0F07")]
    public interface _ControlInReportEvents : _Control
	{

	}
}
