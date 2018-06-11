using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.VBIDEApi
{
    /// <summary>
    /// DispatchInterface _dispVBProjectsEvents
    /// SupportByVersion VBIDE, 12,14,5.3
    /// </summary>
    [SupportByVersion("VBIDE", 12, 14, 5.3)]
    [EntityType(EntityType.IsDispatchInterface)]
	[TypeId("0002E103-0000-0000-C000-000000000046")]
    [CoClassSource(typeof(NetOffice.VBIDEApi._dispVBProjectsEvents))]
    public interface _dispVBProjectsEvents : ICOMObject
    {
        #region Methods

        /// <summary>
        /// SupportByVersion VBIDE 12, 14, 5.3
        /// </summary>
        /// <param name="vBProject">NetOffice.VBIDEApi.VBProject vBProject</param>
        [SupportByVersion("VBIDE", 12, 14, 5.3)]
        void ItemAdded(NetOffice.VBIDEApi.VBProject vBProject);

        /// <summary>
        /// SupportByVersion VBIDE 12, 14, 5.3
        /// </summary>
        /// <param name="vBProject">NetOffice.VBIDEApi.VBProject vBProject</param>
        [SupportByVersion("VBIDE", 12, 14, 5.3)]
        void ItemRemoved(NetOffice.VBIDEApi.VBProject vBProject);

        /// <summary>
        /// SupportByVersion VBIDE 12, 14, 5.3
        /// </summary>
        /// <param name="vBProject">NetOffice.VBIDEApi.VBProject vBProject</param>
        /// <param name="oldName">string oldName</param>
        [SupportByVersion("VBIDE", 12, 14, 5.3)]
        void ItemRenamed(NetOffice.VBIDEApi.VBProject vBProject, string oldName);

        /// <summary>
        /// SupportByVersion VBIDE 12, 14, 5.3
        /// </summary>
        /// <param name="vBProject">NetOffice.VBIDEApi.VBProject vBProject</param>
        [SupportByVersion("VBIDE", 12, 14, 5.3)]
        void ItemActivated(NetOffice.VBIDEApi.VBProject vBProject);

        #endregion
    }
}
