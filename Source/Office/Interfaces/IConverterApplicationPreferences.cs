using System.Reflection;
using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.OfficeApi
{
    /// <summary>
    /// Interface IConverterApplicationPreferences 
    /// SupportByVersion Office, 14,15,16
    /// </summary>
    /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff862807.aspx </remarks>
    [SupportByVersion("Office", 14, 15, 16)]
    [EntityType(EntityType.IsInterface)]
    public interface IConverterApplicationPreferences : ICOMObject
    {
        #region Methods

        /// <summary>
        /// SupportByVersion Office 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff864148.aspx </remarks>
        /// <param name="plcid">Int32 plcid</param>
        [SupportByVersion("Office", 14, 15, 16)]
        Int32 HrGetLcid(out Int32 plcid);

        /// <summary>
        /// SupportByVersion Office 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff860588.aspx </remarks>
        /// <param name="phwnd">Int32 phwnd</param>
        [SupportByVersion("Office", 14, 15, 16)]
        Int32 HrGetHwnd(out Int32 phwnd);

        /// <summary>
        /// SupportByVersion Office 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff864579.aspx </remarks>
        /// <param name="pbstrApplication">string pbstrApplication</param>
        [SupportByVersion("Office", 14, 15, 16)]
        Int32 HrGetApplication(out string pbstrApplication);

        /// <summary>
        /// SupportByVersion Office 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff862557.aspx </remarks>
        /// <param name="pFormat">Int32 pFormat</param>
        [SupportByVersion("Office", 14, 15, 16)]
        Int32 HrCheckFormat(out Int32 pFormat);

        #endregion
    }
}
