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
    /// SupportByVersion VBIDE, 12,14,5.3
    /// </summary>
    [SupportByVersion("VBIDE", 12, 14, 5.3)]
    [EntityType(EntityType.IsCoClass)]
    [ComEventInterface(typeof(EventContracts._dispCommandBarControlEvents))]
    public interface CommandBarEvents : _CommandBarControlEvents, IEventBinding
    {
        #region Events

        /// <summary>
        /// SupportByVersion VBIDE 12 14 5.3
        /// </summary>
        [SupportByVersion("VBIDE", 12, 14, 5.3)]
        event CommandBarEvents_ClickEventHandler ClickEvent;

        #endregion
    }
}
