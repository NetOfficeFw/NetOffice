using System.Reflection;
using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.OfficeApi
{
    /// <summary>
    /// Interface IConverterPreferences 
    /// SupportByVersion Office, 14,15,16
    /// </summary>
    /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff864179.aspx </remarks>
    [SupportByVersion("Office", 14, 15, 16)]
    [EntityType(EntityType.IsInterface)]
	[TypeId("000C03D4-0000-0000-C000-000000000046")]
    public interface IConverterPreferences : ICOMObject
    {
        #region Methods

        /// <summary>
        /// SupportByVersion Office 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff863141.aspx </remarks>
        /// <param name="pfMacroEnabled">Int32 pfMacroEnabled</param>
        [SupportByVersion("Office", 14, 15, 16)]
        Int32 HrGetMacroEnabled(out Int32 pfMacroEnabled);

        /// <summary>
        /// SupportByVersion Office 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff865570.aspx </remarks>
        /// <param name="pFormat">Int32 pFormat</param>
        [SupportByVersion("Office", 14, 15, 16)]
        Int32 HrCheckFormat(out Int32 pFormat);

        /// <summary>
        /// SupportByVersion Office 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff860851.aspx </remarks>
        /// <param name="pfLossySave">Int32 pfLossySave</param>
        [SupportByVersion("Office", 14, 15, 16)]
        Int32 HrGetLossySave(out Int32 pfLossySave);

        #endregion
    }
}
