using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.ADODBApi
{
    /// <summary>
    /// CoClass Command 
    /// SupportByVersion ADODB, 2.1,2.5
    /// </summary>
    [SupportByVersion("ADODB", 2.1, 2.5)]
    [EntityType(EntityType.IsCoClass)]
	[TypeId("00000507-0000-0010-8000-00AA006D2EA4")]
    public interface Command : _Command
    {

    }
}
