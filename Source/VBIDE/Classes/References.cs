using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.VBIDEApi
{
	#region Delegates

	#pragma warning disable
	public delegate void References_ItemAddedEventHandler(NetOffice.VBIDEApi.Reference reference);
	public delegate void References_ItemRemovedEventHandler(NetOffice.VBIDEApi.Reference reference);
    #pragma warning restore

    #endregion

    /// <summary>
    /// CoClass References
    /// SupportByVersion VBIDE, 12,14,5.3
    /// </summary>
    [SupportByVersion("VBIDE", 12, 14, 5.3)]
    [EntityType(EntityType.IsCoClass)]
    [EventSink(typeof(EventInterfaces._dispReferences_Events_SinkHelper))]
    [ComEventInterface(typeof(EventInterfaces._dispReferences_Events))]
    public interface References : _References, IEventBinding
    {
        #region Events

        /// <summary>
        /// SupportByVersion VBIDE 12 14 5.3
        /// </summary>
        [SupportByVersion("VBIDE", 12, 14, 5.3)]
        event References_ItemAddedEventHandler ItemAddedEvent;

        /// <summary>
        /// SupportByVersion VBIDE 12 14 5.3
        /// </summary>
        [SupportByVersion("VBIDE", 12, 14, 5.3)]
        event References_ItemRemovedEventHandler ItemRemovedEvent;

        #endregion
    }
}
