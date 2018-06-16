using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.DAOApi
{
    /// <summary>
    /// CoClass Group 
    /// SupportByVersion DAO, 3.6,12.0
    /// </summary>
    [SupportByVersion("DAO", 3.6, 12.0)]
    [EntityType(EntityType.IsCoClass)]
	[TypeId("00000106-0000-0010-8000-00AA006D2EA4")]
    public interface Group : _Group
    {

    }
}
