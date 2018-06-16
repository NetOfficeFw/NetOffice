using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.ADODBApi
{
    /// <summary>
    /// CoClass Stream 
    /// SupportByVersion ADODB, 2.5
    /// </summary>
    [SupportByVersion("ADODB", 2.5)]
    [EntityType(EntityType.IsCoClass)]
	[TypeId("00000566-0000-0010-8000-00AA006D2EA4")]
    public interface Stream : _Stream
    {

    }
}
