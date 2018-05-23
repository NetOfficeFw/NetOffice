using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.OfficeApi
{
    /// <summary>
    /// DispatchInterface ICustomXMLPartsEvents 
    /// SupportByVersion Office, 12,14,15,16
    /// </summary>
    [SupportByVersion("Office", 12, 14, 15, 16)]
    [EntityType(EntityType.IsDispatchInterface)]
    public interface ICustomXMLPartsEvents : ICOMObject
    {
        #region Methods

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// </summary>
        /// <param name="newPart">NetOffice.OfficeApi.CustomXMLPart newPart</param>
        [SupportByVersion("Office", 12, 14, 15, 16)]
        void PartAfterAdd(NetOffice.OfficeApi.CustomXMLPart newPart);

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// </summary>
        /// <param name="oldPart">NetOffice.OfficeApi.CustomXMLPart oldPart</param>
        [SupportByVersion("Office", 12, 14, 15, 16)]
        void PartBeforeDelete(NetOffice.OfficeApi.CustomXMLPart oldPart);

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// </summary>
        /// <param name="part">NetOffice.OfficeApi.CustomXMLPart part</param>
        [SupportByVersion("Office", 12, 14, 15, 16)]
        void PartAfterLoad(NetOffice.OfficeApi.CustomXMLPart part);

        #endregion
    }
}
