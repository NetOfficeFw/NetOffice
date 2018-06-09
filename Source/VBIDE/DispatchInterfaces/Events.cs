using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.VBIDEApi
{
    /// <summary>
    /// DispatchInterface Events
    /// SupportByVersion VBIDE, 12,14,5.3
    /// </summary>
    [SupportByVersion("VBIDE", 12, 14, 5.3)]
    [EntityType(EntityType.IsDispatchInterface)]
	[TypeId("0002E167-0000-0000-C000-000000000046")]
    public interface Events : ICOMObject
    {
        #region Properties

        /// <summary>
        /// SupportByVersion VBIDE 12, 14, 5.3
        /// Get
        /// </summary>
        /// <param name="vBProject">NetOffice.VBIDEApi.VBProject vBProject</param>
        [SupportByVersion("VBIDE", 12, 14, 5.3)]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        NetOffice.VBIDEApi.ReferencesEvents get_ReferencesEvents(NetOffice.VBIDEApi.VBProject vBProject);

        /// <summary>
        /// SupportByVersion VBIDE 12, 14, 5.3
        /// Alias for get_ReferencesEvents
        /// </summary>
        /// <param name="vBProject">NetOffice.VBIDEApi.VBProject vBProject</param>
        [SupportByVersion("VBIDE", 12, 14, 5.3), Redirect("get_ReferencesEvents")]
        NetOffice.VBIDEApi.ReferencesEvents ReferencesEvents(NetOffice.VBIDEApi.VBProject vBProject);

        /// <summary>
        /// SupportByVersion VBIDE 12, 14, 5.3
        /// Get
        /// </summary>
        /// <param name="commandBarControl">object commandBarControl</param>
        [SupportByVersion("VBIDE", 12, 14, 5.3)]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        NetOffice.VBIDEApi.CommandBarEvents get_CommandBarEvents(object commandBarControl);

        /// <summary>
        /// SupportByVersion VBIDE 12, 14, 5.3
        /// Alias for get_CommandBarEvents
        /// </summary>
        /// <param name="commandBarControl">object commandBarControl</param>
        [SupportByVersion("VBIDE", 12, 14, 5.3), Redirect("get_CommandBarEvents")]
        NetOffice.VBIDEApi.CommandBarEvents CommandBarEvents(object commandBarControl);

        #endregion
    }
}
