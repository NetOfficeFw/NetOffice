using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.OutlookApi
{
	/// <summary>
	/// CoClass _InspectorCtrl 
	/// SupportByVersion Outlook, 10
	/// </summary>
	[SupportByVersion("Outlook", 10)]
	[EntityType(EntityType.IsCoClass)]
    [ComEventContract(typeof(EventContracts._DInspectorEvents))]
	[TypeId("0006F09C-0000-0000-C000-000000000046")]
    public interface _InspectorCtrl : _DInspectorCtrl, IEventBinding
	{

	}
}
