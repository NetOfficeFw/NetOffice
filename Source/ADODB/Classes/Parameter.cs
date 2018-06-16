using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.ADODBApi
{
    /// <summary>
    /// CoClass Parameter 
    /// SupportByVersion ADODB, 2.1,2.5
    /// </summary>
    [SupportByVersion("ADODB", 2.1, 2.5)]
    [EntityType(EntityType.IsCoClass)]
	[TypeId("0000050B-0000-0010-8000-00AA006D2EA4")]
    public interface Parameter : _Parameter
    {

    }
}
