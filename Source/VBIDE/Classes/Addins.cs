using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.VBIDEApi
{
    /// <summary>
	/// CoClass Addins
	/// SupportByVersion VBIDE 12, 14, 5.3
	/// </summary>
	[SupportByVersion("VBIDE", 12, 14, 5.3)]
    [EntityType(EntityType.IsCoClass)]
	[TypeId("DA936B63-AC8B-11D1-B6E5-00A0C90F2744")]
    public interface Addins : NetOffice.VBIDEApi._AddIns
    {

    }
}
