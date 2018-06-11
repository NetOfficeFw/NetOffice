using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.VBIDEApi
{
	#region Delegates

	#pragma warning disable
	public delegate void CommandBarEvents_ClickEventHandler(ICOMObject commandBarControl, ref bool handled, ref bool cancelDefault);
    #pragma warning restore

    #endregion

    /// <summary>
    /// CoClass CommandBarEvents
    /// SupportByVersion VBIDE 12, 14, 5.3
    /// </summary>
    [SupportByVersion("VBIDE", 12, 14, 5.3)]
    [EntityType(EntityType.IsCoClass)]
    [ComEventContract(typeof(EventContracts._dispCommandBarControlEvents))]
	[TypeId("0002E132-0000-0000-C000-000000000046")]
    public interface CommandBarEvents : _CommandBarControlEvents, IEventBinding
    {
        #region Events

        /// <summary>
        /// SupportByVersion VBIDE 12, 14, 5.3
        /// </summary>
        [SupportByVersion("VBIDE", 12, 14, 5.3)]
        event CommandBarEvents_ClickEventHandler ClickEvent;

        #endregion
    }
}
