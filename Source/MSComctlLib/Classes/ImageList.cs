using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.MSComctlLibApi
{
    /// <summary>
    /// CoClass ImageList 
    /// SupportByVersion MSComctlLib, 6
    /// </summary>
    [SupportByVersion("MSComctlLib", 6)]
    [EntityType(EntityType.IsCoClass)]
    [ComEventContract(typeof(EventContracts.ImageListEvents))]
	[TypeId("2C247F23-8591-11D1-B16A-00C0F0283628")]
    public interface ImageList : IImageList, IEventBinding
    {

    }
}
