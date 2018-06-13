using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.OutlookApi
{
	/// <summary>
	/// CoClass _DocSiteControl 
	/// SupportByVersion Outlook, 9,10,11,12,14,15,16
	/// </summary>
	[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
	[EntityType(EntityType.IsCoClass)]
    [ComEventContract(typeof(EventContracts._DDocSiteControlEvents))]
	[TypeId("0006F024-0000-0000-C000-000000000046")]
    public interface _DocSiteControl : _DDocSiteControl, IEventBinding
	{

	}
}
