using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.OfficeApi
{
    /// <summary>
    /// DispatchInterface SharedWorkspaceLink 
    /// SupportByVersion Office, 11,12,14,15,16
    /// </summary>
    /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff865254.aspx </remarks>
    [SupportByVersion("Office", 11, 12, 14, 15, 16)]
    [EntityType(EntityType.IsDispatchInterface)]
    public interface SharedWorkspaceLink : _IMsoDispObj
    {
        #region Properties

        /// <summary>
        /// SupportByVersion Office 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff863119.aspx </remarks>
        [SupportByVersion("Office", 11, 12, 14, 15, 16)]
        string URL { get; set; }

        /// <summary>
        /// SupportByVersion Office 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff860499.aspx </remarks>
        [SupportByVersion("Office", 11, 12, 14, 15, 16)]
        string Description { get; set; }

        /// <summary>
        /// SupportByVersion Office 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff861886.aspx </remarks>
        [SupportByVersion("Office", 11, 12, 14, 15, 16)]
        string Notes { get; set; }

        /// <summary>
        /// SupportByVersion Office 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff863662.aspx </remarks>
        [SupportByVersion("Office", 11, 12, 14, 15, 16)]
        string CreatedBy { get; }

        /// <summary>
        /// SupportByVersion Office 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff861532.aspx </remarks>
        [SupportByVersion("Office", 11, 12, 14, 15, 16)]
        object CreatedDate { get; }

        /// <summary>
        /// SupportByVersion Office 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff861094.aspx </remarks>
        [SupportByVersion("Office", 11, 12, 14, 15, 16)]
        string ModifiedBy { get; }

        /// <summary>
        /// SupportByVersion Office 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff860294.aspx </remarks>
        [SupportByVersion("Office", 11, 12, 14, 15, 16)]
        object ModifiedDate { get; }

        /// <summary>
        /// SupportByVersion Office 11, 12, 14, 15, 16
        /// Get
        /// Unknown COM Proxy
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff863516.aspx </remarks>
        [SupportByVersion("Office", 11, 12, 14, 15, 16), ProxyResult]
        object Parent { get; }

        #endregion

        #region Methods

        /// <summary>
        /// SupportByVersion Office 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff862046.aspx </remarks>
        [SupportByVersion("Office", 11, 12, 14, 15, 16)]
        void Save();

        /// <summary>
        /// SupportByVersion Office 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff863052.aspx </remarks>
        [SupportByVersion("Office", 11, 12, 14, 15, 16)]
        void Delete();

        #endregion
    }
}
