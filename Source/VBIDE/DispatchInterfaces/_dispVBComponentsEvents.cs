using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.VBIDEApi
{
    /// <summary>
    /// DispatchInterface _dispVBComponentsEvents
    /// SupportByVersion VBIDE, 12,14,5.3
    /// </summary>
	[TypeId("0002E116-0000-0000-C000-000000000046")]
    public interface _dispVBComponentsEvents : ICOMObject
    {
        #region Methods

        /// <summary>
        /// SupportByVersion VBIDE 12, 14, 5.3
        /// </summary>
        /// <param name="vBComponent">NetOffice.VBIDEApi.VBComponent vBComponent</param>
        [SupportByVersion("VBIDE", 12, 14, 5.3)]
        void ItemAdded(NetOffice.VBIDEApi.VBComponent vBComponent);

        /// <summary>
        /// SupportByVersion VBIDE 12, 14, 5.3
        /// </summary>
        /// <param name="vBComponent">NetOffice.VBIDEApi.VBComponent vBComponent</param>
        [SupportByVersion("VBIDE", 12, 14, 5.3)]
        void ItemRemoved(NetOffice.VBIDEApi.VBComponent vBComponent);

        /// <summary>
        /// SupportByVersion VBIDE 12, 14, 5.3
        /// </summary>
        /// <param name="vBComponent">NetOffice.VBIDEApi.VBComponent vBComponent</param>
        /// <param name="oldName">string oldName</param>
        [SupportByVersion("VBIDE", 12, 14, 5.3)]
        void ItemRenamed(NetOffice.VBIDEApi.VBComponent vBComponent, string oldName);

        /// <summary>
        /// SupportByVersion VBIDE 12, 14, 5.3
        /// </summary>
        /// <param name="vBComponent">NetOffice.VBIDEApi.VBComponent vBComponent</param>
        [SupportByVersion("VBIDE", 12, 14, 5.3)]
        void ItemSelected(NetOffice.VBIDEApi.VBComponent vBComponent);

        /// <summary>
        /// SupportByVersion VBIDE 12, 14, 5.3
        /// </summary>
        /// <param name="vBComponent">NetOffice.VBIDEApi.VBComponent vBComponent</param>
        [SupportByVersion("VBIDE", 12, 14, 5.3)]
        void ItemActivated(NetOffice.VBIDEApi.VBComponent vBComponent);

        /// <summary>
        /// SupportByVersion VBIDE 12, 14, 5.3
        /// </summary>
        /// <param name="vBComponent">NetOffice.VBIDEApi.VBComponent vBComponent</param>
        [SupportByVersion("VBIDE", 12, 14, 5.3)]
        void ItemReloaded(NetOffice.VBIDEApi.VBComponent vBComponent);

        #endregion
    }
}
