using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.MSComctlLibApi
{
    /// <summary>
    /// CoClass DataObject 
    /// SupportByVersion MSComctlLib, 6
    /// </summary>
    [SupportByVersion("MSComctlLib", 6)]
    [EntityType(EntityType.IsCoClass)]
	[TypeId("2334D2B2-713E-11CF-8AE5-00AA00C00905")]
    public interface DataObject : IVBDataObject
    {

    }
}
