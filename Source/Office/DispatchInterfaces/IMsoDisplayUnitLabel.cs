using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.OfficeApi
{
    /// <summary>
    /// DispatchInterface IMsoDisplayUnitLabel 
    /// SupportByVersion Office, 12,14,15,16
    /// </summary>
    [SupportByVersion("Office", 12, 14, 15, 16)]
    [EntityType(EntityType.IsDispatchInterface)]
	[TypeId("6EA00553-9439-4D5A-B1E6-DC15A54DA8B2")]
    public interface IMsoDisplayUnitLabel : IMsoChartTitle
    {

    }
}
