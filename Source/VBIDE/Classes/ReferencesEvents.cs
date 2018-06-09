using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.VBIDEApi
{
	#region Delegates

	#pragma warning disable
	public delegate void ReferencesEvents_ItemAddedEventHandler(NetOffice.VBIDEApi.Reference reference);
	public delegate void ReferencesEvents_ItemRemovedEventHandler(NetOffice.VBIDEApi.Reference reference);
    #pragma warning restore

    #endregion

    /// <summary>
    /// CoClass ReferencesEvents
    /// SupportByVersion VBIDE, 12,14,5.3
    /// </summary>
    [SupportByVersion("VBIDE", 12, 14, 5.3)]
    [EntityType(EntityType.IsCoClass)]
    [ComEventContract(typeof(EventContracts._dispReferencesEvents))]
	[TypeId("0002E119-0000-0000-C000-000000000046")]
    public interface ReferencesEvents : _ReferencesEvents, IEventBinding
    {
        #region Events

        /// <summary>
        /// SupportByVersion VBIDE 12 14 5.3
        /// </summary>
        [SupportByVersion("VBIDE", 12, 14, 5.3)]
        event ReferencesEvents_ItemAddedEventHandler ItemAddedEvent;

        /// <summary>
        /// SupportByVersion VBIDE 12 14 5.3
        /// </summary>
        [SupportByVersion("VBIDE", 12, 14, 5.3)]
        event ReferencesEvents_ItemRemovedEventHandler ItemRemovedEvent;

        #endregion
    }
}
