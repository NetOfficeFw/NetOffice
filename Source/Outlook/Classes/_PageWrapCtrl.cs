using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.OutlookApi
{
	/// <summary>
	/// CoClass _PageWrapCtrl 
	/// SupportByVersion Outlook, 10
	/// </summary>
	[SupportByVersion("Outlook", 10)]
	[EntityType(EntityType.IsCoClass)]
    [ComEventContract(typeof(EventContracts._DPageWrapCtrlEvents))]
	[TypeId("0006F098-0000-0000-C000-000000000046")]
    public interface _PageWrapCtrl : _DPageWrapCtrl, IEventBinding
	{

	}
}
