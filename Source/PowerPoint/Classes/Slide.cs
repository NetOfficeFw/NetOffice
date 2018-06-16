using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.PowerPointApi
{
    /// <summary>
    /// CoClass Slide
    /// SupportByVersion PowerPoint, 9,10,11,12,14,15,16
    /// </summary>
    /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff745958.aspx </remarks>
    [SupportByVersion("PowerPoint", 9, 10, 11, 12, 14, 15, 16)]
    [EntityType(EntityType.IsCoClass)]
    [ComEventContract(typeof(EventContracts.SldEvents))]
	[TypeId("91493445-5A91-11CF-8700-00AA0060263B")]    
    public interface Slide : _Slide, IEventBinding
    {

    }
}
