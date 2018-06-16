using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.OWC10Api
{
	#region Delegates

	#pragma warning disable
	public delegate void Range_ChangeEventHandler();
    #pragma warning restore

    #endregion

    /// <summary>
    /// CoClass Range 
    /// SupportByVersion OWC10, 1
    /// </summary>
    [SupportByVersion("OWC10", 1)]
    [EntityType(EntityType.IsCoClass)]
    [ComEventContract(typeof(EventContracts.IRangeEvents))]
	[TypeId("19A4E1A0-9334-4EB0-BD78-0AE080B8B4A7")]
    public interface Range : _Range, IEventBinding
    {
        #region Events

        /// <summary>
        /// SupportByVersion OWC10 1
        /// </summary>
        [SupportByVersion("OWC10", 1)]
        event Range_ChangeEventHandler ChangeEvent;

        #endregion
    }
}
