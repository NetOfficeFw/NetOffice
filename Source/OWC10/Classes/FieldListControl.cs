using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.OWC10Api
{
    /// <summary>
    /// CoClass FieldListControl 
    /// SupportByVersion OWC10, 1
    /// </summary>
    [SupportByVersion("OWC10", 1)]
    [EntityType(EntityType.IsCoClass)]
	[TypeId("0002E557-0000-0000-C000-000000000046")]
    public interface FieldListControl : FieldList
    {

    }
}
