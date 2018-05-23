using System.Reflection;
using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.OfficeApi
{
    /// <summary>
    /// DispatchInterface DocumentInspector 
    /// SupportByVersion Office, 12,14,15,16
    /// </summary>
    /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff862517.aspx </remarks>
    [SupportByVersion("Office", 12, 14, 15, 16)]
    [EntityType(EntityType.IsDispatchInterface)]
    public interface DocumentInspector : _IMsoDispObj
    {
        #region Properties

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff862757.aspx </remarks>
        [SupportByVersion("Office", 12, 14, 15, 16)]
        string Name { get; }

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff860548.aspx </remarks>
        [SupportByVersion("Office", 12, 14, 15, 16)]
        string Description { get; }

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// Get
        /// Unknown COM Proxy
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff863644.aspx </remarks>
        [SupportByVersion("Office", 12, 14, 15, 16), ProxyResult]
        object Parent { get; }

        #endregion

        #region Methods

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff861849.aspx </remarks>
        /// <param name="status">NetOffice.OfficeApi.Enums.MsoDocInspectorStatus status</param>
        /// <param name="results">string results</param>
        [SupportByVersion("Office", 12, 14, 15, 16)]
        void Inspect(out NetOffice.OfficeApi.Enums.MsoDocInspectorStatus status, out string results);

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff863804.aspx </remarks>
        /// <param name="status">NetOffice.OfficeApi.Enums.MsoDocInspectorStatus status</param>
        /// <param name="results">string results</param>
        [SupportByVersion("Office", 12, 14, 15, 16)]
        void Fix(out NetOffice.OfficeApi.Enums.MsoDocInspectorStatus status, out string results);

        #endregion
    }
}
