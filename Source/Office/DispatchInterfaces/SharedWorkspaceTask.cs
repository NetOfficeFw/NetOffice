using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.OfficeApi
{
    /// <summary>
    /// DispatchInterface SharedWorkspaceTask 
    /// SupportByVersion Office, 11,12,14,15,16
    /// </summary>
    /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff865531.aspx </remarks>
    [SupportByVersion("Office", 11, 12, 14, 15, 16)]
    [EntityType(EntityType.IsDispatchInterface)]
    public interface SharedWorkspaceTask : _IMsoDispObj
    {
        #region Properties

        /// <summary>
        /// SupportByVersion Office 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff860234.aspx </remarks>
        [SupportByVersion("Office", 11, 12, 14, 15, 16)]
        string Title { get; set; }

        /// <summary>
        /// SupportByVersion Office 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff861531.aspx </remarks>
        [SupportByVersion("Office", 11, 12, 14, 15, 16)]
        string AssignedTo { get; set; }

        /// <summary>
        /// SupportByVersion Office 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff864957.aspx </remarks>
        [SupportByVersion("Office", 11, 12, 14, 15, 16)]
        NetOffice.OfficeApi.Enums.MsoSharedWorkspaceTaskStatus Status { get; set; }

        /// <summary>
        /// SupportByVersion Office 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff863054.aspx </remarks>
        [SupportByVersion("Office", 11, 12, 14, 15, 16)]
        NetOffice.OfficeApi.Enums.MsoSharedWorkspaceTaskPriority Priority { get; set; }

        /// <summary>
        /// SupportByVersion Office 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff860514.aspx </remarks>
        [SupportByVersion("Office", 11, 12, 14, 15, 16)]
        string Description { get; set; }

        /// <summary>
        /// SupportByVersion Office 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff862835.aspx </remarks>
        [SupportByVersion("Office", 11, 12, 14, 15, 16)]
        object DueDate { get; set; }

        /// <summary>
        /// SupportByVersion Office 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff864667.aspx </remarks>
        [SupportByVersion("Office", 11, 12, 14, 15, 16)]
        string CreatedBy { get; }

        /// <summary>
        /// SupportByVersion Office 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff862213.aspx </remarks>
        [SupportByVersion("Office", 11, 12, 14, 15, 16)]
        object CreatedDate { get; }

        /// <summary>
        /// SupportByVersion Office 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff864980.aspx </remarks>
        [SupportByVersion("Office", 11, 12, 14, 15, 16)]
        string ModifiedBy { get; }

        /// <summary>
        /// SupportByVersion Office 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff860842.aspx </remarks>
        [SupportByVersion("Office", 11, 12, 14, 15, 16)]
        object ModifiedDate { get; }

        /// <summary>
        /// SupportByVersion Office 11, 12, 14, 15, 16
        /// Get
        /// Unknown COM Proxy
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff862819.aspx </remarks>
        [SupportByVersion("Office", 11, 12, 14, 15, 16), ProxyResult]
        object Parent { get; }

        #endregion

        #region Methods

        /// <summary>
        /// SupportByVersion Office 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff865262.aspx </remarks>
        [SupportByVersion("Office", 11, 12, 14, 15, 16)]
        void Save();

        /// <summary>
        /// SupportByVersion Office 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff862097.aspx </remarks>
        [SupportByVersion("Office", 11, 12, 14, 15, 16)]
        void Delete();

        #endregion
    }
}
