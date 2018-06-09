using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.OfficeApi
{
    /// <summary>
    /// DispatchInterface ICommandBarButtonEvents 
    /// SupportByVersion Office, 9,10,11,12,14,15,16
    /// </summary>
    [SupportByVersion("Office", 9, 10, 11, 12, 14, 15, 16)]
    [EntityType(EntityType.IsDispatchInterface)]
	[TypeId("55F88890-7708-11D1-ACEB-006008961DA5")]
    public interface ICommandBarButtonEvents : ICOMObject
    {
        #region Methods

        /// <summary>
        /// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="ctrl">NetOffice.OfficeApi.CommandBarButton ctrl</param>
        /// <param name="cancelDefault">bool cancelDefault</param>
        [SupportByVersion("Office", 9, 10, 11, 12, 14, 15, 16)]
        void Click(NetOffice.OfficeApi.CommandBarButton ctrl, bool cancelDefault);

        #endregion
    }
}
