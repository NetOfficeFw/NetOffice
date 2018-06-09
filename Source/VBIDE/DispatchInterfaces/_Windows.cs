using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.VBIDEApi
{
    /// <summary>
    /// DispatchInterface _Windows
    /// SupportByVersion VBIDE, 12,14,5.3
    /// </summary>
    [SupportByVersion("VBIDE", 12, 14, 5.3)]
    [EntityType(EntityType.IsDispatchInterface), BaseType]
	[TypeId("F57B7ED0-D8AB-11D1-85DF-00C04F98F42C")]
    public interface _Windows : _Windows_old
    {
        #region Methods

        /// <summary>
        /// SupportByVersion VBIDE 12, 14, 5.3
        /// </summary>
        /// <param name="addInInst">NetOffice.VBIDEApi.AddIn addInInst</param>
        /// <param name="progId">string progId</param>
        /// <param name="caption">string caption</param>
        /// <param name="guidPosition">string guidPosition</param>
        /// <param name="docObj">object docObj</param>
        [SupportByVersion("VBIDE", 12, 14, 5.3)]
        NetOffice.VBIDEApi.Window CreateToolWindow(NetOffice.VBIDEApi.AddIn addInInst, string progId, string caption, string guidPosition, object docObj);

        #endregion
    }
}
