using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.VBIDEApi
{
    /// <summary>
    /// CoClass CodeModule
    /// SupportByVersion VBIDE 12, 14, 5.3
    /// </summary>
    [SupportByVersion("VBIDE", 12, 14, 5.3)]
    [EntityType(EntityType.IsCoClass)]
	[TypeId("0002E170-0000-0000-C000-000000000046")]
    public interface CodeModule : _CodeModule
    {

    }
}
