using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.OutlookApi
{
	/// <summary>
	/// CoClass _RecipientControl 
	/// SupportByVersion Outlook, 9,10,11,12,14,15,16
	/// </summary>
	[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
	[EntityType(EntityType.IsCoClass)]
    [ComEventContract(typeof(EventContracts._DRecipientControlEvents))]
	[TypeId("0006F023-0000-0000-C000-000000000046")]   
    public interface _RecipientControl : _DRecipientControl, IEventBinding
	{

	}
}
