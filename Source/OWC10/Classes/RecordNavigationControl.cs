using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.OWC10Api
{
	#region Delegates

	#pragma warning disable
	public delegate void RecordNavigationControl_ButtonClickEventHandler(NetOffice.OWC10Api.Enums.NavButtonEnum navButton);
    #pragma warning restore

    #endregion

    /// <summary>
    /// CoClass RecordNavigationControl 
    /// SupportByVersion OWC10, 1
    /// </summary>
    [SupportByVersion("OWC10", 1)]
    [EntityType(EntityType.IsCoClass)]
    [ComEventContract(typeof(EventContracts._NavigationEvent))]
	[TypeId("0002E554-0000-0000-C000-000000000046")]
    public interface RecordNavigationControl : INavigationControl, IEventBinding
    {
        #region Events

        /// <summary>
        /// SupportByVersion OWC10 1
        /// </summary>
        [SupportByVersion("OWC10", 1)]
        event RecordNavigationControl_ButtonClickEventHandler ButtonClickEvent;

        #endregion
    }
}
