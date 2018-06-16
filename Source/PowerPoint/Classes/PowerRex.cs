using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.PowerPointApi
{
    /// <summary>
    /// CoClass PowerRex
    /// SupportByVersion PowerPoint, 10,11,12,14,15,16
    /// </summary>
    [SupportByVersion("PowerPoint", 10, 11, 12, 14, 15, 16)]
    [EntityType(EntityType.IsCoClass)]
	[TypeId("91493448-5A91-11CF-8700-00AA0060263B")]
    public interface PowerRex : _PowerRex
    {

    }
}
