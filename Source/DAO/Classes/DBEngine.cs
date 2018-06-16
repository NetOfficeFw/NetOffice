using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.DAOApi
{
    /// <summary>
    /// CoClass DBEngine 
    /// SupportByVersion DAO, 3.6,12.0
    /// </summary>
    [SupportByVersion("DAO", 3.6, 12.0)]
    [EntityType(EntityType.IsCoClass)]
	[TypeId("00000100-0000-0010-8000-00AA006D2EA4")]
    public interface DBEngine : _DBEngine
    {

    }
}
