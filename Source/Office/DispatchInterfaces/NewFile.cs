using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.OfficeApi
{
    /// <summary>
    /// DispatchInterface NewFile 
    /// SupportByVersion Office, 10,11,12,14,15,16
    /// </summary>
    /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff862417.aspx </remarks>
    [SupportByVersion("Office", 10, 11, 12, 14, 15, 16)]
    [EntityType(EntityType.IsDispatchInterface)]
	[TypeId("000C0936-0000-0000-C000-000000000046")]
    public interface NewFile : _IMsoDispObj
    {
        #region Methods

        /// <summary>
        /// SupportByVersion Office 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff860279.aspx </remarks>
        /// <param name="fileName">string fileName</param>
        /// <param name="section">optional object section</param>
        /// <param name="displayName">optional object displayName</param>
        /// <param name="action">optional object action</param>
        [SupportByVersion("Office", 10, 11, 12, 14, 15, 16)]
        bool Add(string fileName, object section, object displayName, object action);

        /// <summary>
        /// SupportByVersion Office 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff860279.aspx </remarks>
        /// <param name="fileName">string fileName</param>
        [CustomMethod]
        [SupportByVersion("Office", 10, 11, 12, 14, 15, 16)]
        bool Add(string fileName);

        /// <summary>
        /// SupportByVersion Office 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff860279.aspx </remarks>
        /// <param name="fileName">string fileName</param>
        /// <param name="section">optional object section</param>
        [CustomMethod]
        [SupportByVersion("Office", 10, 11, 12, 14, 15, 16)]
        bool Add(string fileName, object section);

        /// <summary>
        /// SupportByVersion Office 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff860279.aspx </remarks>
        /// <param name="fileName">string fileName</param>
        /// <param name="section">optional object section</param>
        /// <param name="displayName">optional object displayName</param>
        [CustomMethod]
        [SupportByVersion("Office", 10, 11, 12, 14, 15, 16)]
        bool Add(string fileName, object section, object displayName);

        /// <summary>
        /// SupportByVersion Office 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff860573.aspx </remarks>
        /// <param name="fileName">string fileName</param>
        /// <param name="section">optional object section</param>
        /// <param name="displayName">optional object displayName</param>
        /// <param name="action">optional object action</param>
        [SupportByVersion("Office", 10, 11, 12, 14, 15, 16)]
        bool Remove(string fileName, object section, object displayName, object action);

        /// <summary>
        /// SupportByVersion Office 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff860573.aspx </remarks>
        /// <param name="fileName">string fileName</param>
        [CustomMethod]
        [SupportByVersion("Office", 10, 11, 12, 14, 15, 16)]
        bool Remove(string fileName);

        /// <summary>
        /// SupportByVersion Office 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff860573.aspx </remarks>
        /// <param name="fileName">string fileName</param>
        /// <param name="section">optional object section</param>
        [CustomMethod]
        [SupportByVersion("Office", 10, 11, 12, 14, 15, 16)]
        bool Remove(string fileName, object section);

        /// <summary>
        /// SupportByVersion Office 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff860573.aspx </remarks>
        /// <param name="fileName">string fileName</param>
        /// <param name="section">optional object section</param>
        /// <param name="displayName">optional object displayName</param>
        [CustomMethod]
        [SupportByVersion("Office", 10, 11, 12, 14, 15, 16)]
        bool Remove(string fileName, object section, object displayName);

        #endregion
    }
}
