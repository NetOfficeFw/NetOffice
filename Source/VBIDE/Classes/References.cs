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
    [ComEventContract(typeof(EventContracts._dispReferences_Events))]
	[TypeId("0002E17C-0000-0000-C000-000000000046")]
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
