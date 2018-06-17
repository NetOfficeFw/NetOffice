using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.AccessApi
{
	/// <summary>
	/// CoClass _ChildLabel
	/// SupportByVersion Access, 9,10,11,12,14,15,16
	/// </summary>
	[SupportByVersion("Access", 9,10,11,12,14,15,16)]
	[EntityType(EntityType.IsCoClass)]
    [ComEventContract(typeof(EventContracts._ChildLabelEvents), typeof(EventContracts.DispChildLabelEvents))]
	[TypeId("BC9E4359-F037-11CD-8701-00AA003F0F07")]
    public interface _ChildLabel : _Label
	{

	}
}
