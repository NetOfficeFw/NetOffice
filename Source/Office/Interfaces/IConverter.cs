using System.Reflection;
using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.OfficeApi
{
    /// <summary>
    /// Interface IConverter 
    /// SupportByVersion Office, 14,15,16
    /// </summary>
    /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff861235.aspx </remarks>
    [SupportByVersion("Office", 14, 15, 16)]
    [EntityType(EntityType.IsInterface)]
	[TypeId("000C03D7-0000-0000-C000-000000000046")]
    public interface IConverter : ICOMObject
    {
        #region Methods

        /// <summary>
        /// SupportByVersion Office 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff864088.aspx </remarks>
        /// <param name="pcap">NetOffice.OfficeApi.IConverterApplicationPreferences pcap</param>
        /// <param name="ppcp">NetOffice.OfficeApi.IConverterPreferences ppcp</param>
        /// <param name="pcuic">NetOffice.OfficeApi.IConverterUICallback pcuic</param>
        [SupportByVersion("Office", 14, 15, 16)]
        Int32 HrInitConverter(NetOffice.OfficeApi.IConverterApplicationPreferences pcap, out NetOffice.OfficeApi.IConverterPreferences ppcp, NetOffice.OfficeApi.IConverterUICallback pcuic);

        /// <summary>
        /// SupportByVersion Office 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff862058.aspx </remarks>
        /// <param name="pcuic">NetOffice.OfficeApi.IConverterUICallback pcuic</param>
        [SupportByVersion("Office", 14, 15, 16)]
        Int32 HrUninitConverter(NetOffice.OfficeApi.IConverterUICallback pcuic);

        /// <summary>
        /// SupportByVersion Office 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff864636.aspx </remarks>
        /// <param name="bstrSourcePath">string bstrSourcePath</param>
        /// <param name="bstrDestPath">string bstrDestPath</param>
        /// <param name="pcap">NetOffice.OfficeApi.IConverterApplicationPreferences pcap</param>
        /// <param name="ppcp">NetOffice.OfficeApi.IConverterPreferences ppcp</param>
        /// <param name="pcuic">NetOffice.OfficeApi.IConverterUICallback pcuic</param>
        [SupportByVersion("Office", 14, 15, 16)]
        Int32 HrImport(string bstrSourcePath, string bstrDestPath, NetOffice.OfficeApi.IConverterApplicationPreferences pcap, out NetOffice.OfficeApi.IConverterPreferences ppcp, NetOffice.OfficeApi.IConverterUICallback pcuic);

        /// <summary>
        /// SupportByVersion Office 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff863696.aspx </remarks>
        /// <param name="bstrSourcePath">string bstrSourcePath</param>
        /// <param name="bstrDestPath">string bstrDestPath</param>
        /// <param name="bstrClass">string bstrClass</param>
        /// <param name="pcap">NetOffice.OfficeApi.IConverterApplicationPreferences pcap</param>
        /// <param name="ppcp">NetOffice.OfficeApi.IConverterPreferences ppcp</param>
        /// <param name="pcuic">NetOffice.OfficeApi.IConverterUICallback pcuic</param>
        [SupportByVersion("Office", 14, 15, 16)]
        Int32 HrExport(string bstrSourcePath, string bstrDestPath, string bstrClass, NetOffice.OfficeApi.IConverterApplicationPreferences pcap, out NetOffice.OfficeApi.IConverterPreferences ppcp, NetOffice.OfficeApi.IConverterUICallback pcuic);

        /// <summary>
        /// SupportByVersion Office 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff864094.aspx </remarks>
        /// <param name="bstrPath">string bstrPath</param>
        /// <param name="pbstrClass">string pbstrClass</param>
        /// <param name="pcap">NetOffice.OfficeApi.IConverterApplicationPreferences pcap</param>
        /// <param name="ppcp">NetOffice.OfficeApi.IConverterPreferences ppcp</param>
        /// <param name="pcuic">NetOffice.OfficeApi.IConverterUICallback pcuic</param>
        [SupportByVersion("Office", 14, 15, 16)]
        Int32 HrGetFormat(string bstrPath, out string pbstrClass, NetOffice.OfficeApi.IConverterApplicationPreferences pcap, out NetOffice.OfficeApi.IConverterPreferences ppcp, NetOffice.OfficeApi.IConverterUICallback pcuic);

        /// <summary>
        /// SupportByVersion Office 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff861511.aspx </remarks>
        /// <param name="hrErr">Int32 hrErr</param>
        /// <param name="pbstrErrorMsg">string pbstrErrorMsg</param>
        /// <param name="pcap">NetOffice.OfficeApi.IConverterApplicationPreferences pcap</param>
        [SupportByVersion("Office", 14, 15, 16)]
        Int32 HrGetErrorString(Int32 hrErr, out string pbstrErrorMsg, NetOffice.OfficeApi.IConverterApplicationPreferences pcap);

        #endregion
    }
}
