using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.PowerPointApi
{
    /// <summary>
	/// CoClass Master
	/// SupportByVersion PowerPoint, 9,10,11,12,14,15,16
	/// </summary>
	/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff744232.aspx </remarks>
	[SupportByVersion("PowerPoint", 9, 10, 11, 12, 14, 15, 16)]
    [EntityType(EntityType.IsCoClass)]
    [ComEventContract(typeof(EventContracts.MasterEvents))]
	[TypeId("91493447-5A91-11CF-8700-00AA0060263B")]
    public interface Master : _Master, IEventBinding
    {

    }
}
