using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.OWC10Api
{
    /// <summary>
    /// CoClass OfflineInfo 
    /// SupportByVersion OWC10, 1
    /// </summary>
    [SupportByVersion("OWC10", 1)]
    [EntityType(EntityType.IsCoClass)]
	[TypeId("E2AC0C6A-7079-11D3-8D01-0050048383A8")]
    public interface OfflineInfo : IOfflineInfo
    {

    }
}
