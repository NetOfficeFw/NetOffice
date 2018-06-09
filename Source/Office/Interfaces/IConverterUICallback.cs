using System.Reflection;
using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.OfficeApi
{
    /// <summary>
    /// Interface IConverterUICallback 
    /// SupportByVersion Office, 14,15,16
    /// </summary>
    /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff863370.aspx </remarks>
    [SupportByVersion("Office", 14, 15, 16)]
    [EntityType(EntityType.IsInterface)]
	[TypeId("000C03D6-0000-0000-C000-000000000046")]
    public interface IConverterUICallback : ICOMObject
    {
        #region Methods

        /// <summary>
        /// SupportByVersion Office 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff861826.aspx </remarks>
        /// <param name="uPercentComplete">UIntPtr uPercentComplete</param>
        [SupportByVersion("Office", 14, 15, 16)]
        Int32 HrReportProgress(UIntPtr uPercentComplete);

        /// <summary>
        /// SupportByVersion Office 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff861376.aspx </remarks>
        /// <param name="bstrText">string bstrText</param>
        /// <param name="bstrCaption">string bstrCaption</param>
        /// <param name="uType">UIntPtr uType</param>
        /// <param name="pidResult">Int32 pidResult</param>
        [SupportByVersion("Office", 14, 15, 16)]
        Int32 HrMessageBox(string bstrText, string bstrCaption, UIntPtr uType, out Int32 pidResult);

        /// <summary>
        /// SupportByVersion Office 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff861803.aspx </remarks>
        /// <param name="bstrText">string bstrText</param>
        /// <param name="bstrCaption">string bstrCaption</param>
        /// <param name="pbstrInput">string pbstrInput</param>
        /// <param name="fPassword">Int32 fPassword</param>
        [SupportByVersion("Office", 14, 15, 16)]
        Int32 HrInputBox(string bstrText, string bstrCaption, out string pbstrInput, Int32 fPassword);

        #endregion
    }
}
