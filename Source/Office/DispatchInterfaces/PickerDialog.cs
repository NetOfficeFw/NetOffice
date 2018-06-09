using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.OfficeApi
{
    /// <summary>
    /// DispatchInterface PickerDialog 
    /// SupportByVersion Office, 14,15,16
    /// </summary>
    /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff860858.aspx </remarks>
    [SupportByVersion("Office", 14, 15, 16)]
    [EntityType(EntityType.IsDispatchInterface)]
	[TypeId("000C03E6-0000-0000-C000-000000000046")]
    public interface PickerDialog : _IMsoDispObj
    {
        #region Properties

        /// <summary>
        /// SupportByVersion Office 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff862371.aspx </remarks>
        [SupportByVersion("Office", 14, 15, 16)]
        string DataHandlerId { get; set; }

        /// <summary>
        /// SupportByVersion Office 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff862526.aspx </remarks>
        [SupportByVersion("Office", 14, 15, 16)]
        string Title { get; set; }

        /// <summary>
        /// SupportByVersion Office 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff860248.aspx </remarks>
        [SupportByVersion("Office", 14, 15, 16)]
        NetOffice.OfficeApi.PickerProperties Properties { get; }

        #endregion

        #region Methods

        /// <summary>
        /// SupportByVersion Office 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff861181.aspx </remarks>
        [SupportByVersion("Office", 14, 15, 16)]
        NetOffice.OfficeApi.PickerResults CreatePickerResults();

        /// <summary>
        /// SupportByVersion Office 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff861095.aspx </remarks>
        /// <param name="isMultiSelect">optional bool IsMultiSelect = true</param>
        /// <param name="existingResults">optional NetOffice.OfficeApi.PickerResults ExistingResults = 0</param>
        [SupportByVersion("Office", 14, 15, 16)]
        NetOffice.OfficeApi.PickerResults Show(object isMultiSelect, object existingResults);

        /// <summary>
        /// SupportByVersion Office 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff861095.aspx </remarks>
        [CustomMethod]
        [SupportByVersion("Office", 14, 15, 16)]
        NetOffice.OfficeApi.PickerResults Show();

        /// <summary>
        /// SupportByVersion Office 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff861095.aspx </remarks>
        /// <param name="isMultiSelect">optional bool IsMultiSelect = true</param>
        [CustomMethod]
        [SupportByVersion("Office", 14, 15, 16)]
        NetOffice.OfficeApi.PickerResults Show(object isMultiSelect);

        /// <summary>
        /// SupportByVersion Office 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff861733.aspx </remarks>
        /// <param name="tokenText">string tokenText</param>
        /// <param name="duplicateDlgMode">Int32 duplicateDlgMode</param>
        [SupportByVersion("Office", 14, 15, 16)]
        NetOffice.OfficeApi.PickerResults Resolve(string tokenText, Int32 duplicateDlgMode);

        #endregion
    }
}
