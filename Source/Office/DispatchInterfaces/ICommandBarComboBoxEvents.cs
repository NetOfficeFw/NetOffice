using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.OfficeApi
{
    /// <summary>
    /// DispatchInterface ICommandBarComboBoxEvents 
    /// SupportByVersion Office, 9,10,11,12,14,15,16
    /// </summary>
    [SupportByVersion("Office", 9, 10, 11, 12, 14, 15, 16)]
    [EntityType(EntityType.IsDispatchInterface)]
    public interface ICommandBarComboBoxEvents : ICOMObject
    {
        #region Methods

        /// <summary>
        /// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="ctrl">NetOffice.OfficeApi.CommandBarComboBox ctrl</param>
        [SupportByVersion("Office", 9, 10, 11, 12, 14, 15, 16)]
        void Change(NetOffice.OfficeApi.CommandBarComboBox ctrl);

        #endregion
    }
}
