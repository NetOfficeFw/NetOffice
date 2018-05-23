using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.OfficeApi
{
    /// <summary>
    /// DispatchInterface IMsoDispCagNotifySink 
    /// SupportByVersion Office, 9,10,11,12,14,15,16
    /// </summary>
    [SupportByVersion("Office", 9, 10, 11, 12, 14, 15, 16)]
    [EntityType(EntityType.IsDispatchInterface)]
    public interface IMsoDispCagNotifySink : ICOMObject
    {
        #region Methods

        /// <summary>
        /// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="pClipMoniker">object pClipMoniker</param>
        /// <param name="pItemMoniker">object pItemMoniker</param>
        [SupportByVersion("Office", 9, 10, 11, 12, 14, 15, 16)]
        void InsertClip(object pClipMoniker, object pItemMoniker);

        /// <summary>
        /// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        [SupportByVersion("Office", 9, 10, 11, 12, 14, 15, 16)]
        void WindowIsClosing();

        #endregion
    }
}
