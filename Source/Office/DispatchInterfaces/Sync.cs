using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.OfficeApi
{
    /// <summary>
    /// DispatchInterface Sync 
    /// SupportByVersion Office, 11,12,14,15,16
    /// </summary>
    /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff860602.aspx </remarks>
    [SupportByVersion("Office", 11, 12, 14, 15, 16)]
    [EntityType(EntityType.IsDispatchInterface)]
	[TypeId("000C0386-0000-0000-C000-000000000046")]
    public interface Sync : _IMsoDispObj
    {
        #region Properties

        /// <summary>
        /// SupportByVersion Office 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff865564.aspx </remarks>
        [SupportByVersion("Office", 11, 12, 14, 15, 16)]
        NetOffice.OfficeApi.Enums.MsoSyncStatusType Status { get; }

        /// <summary>
        /// SupportByVersion Office 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff865364.aspx </remarks>
        [SupportByVersion("Office", 11, 12, 14, 15, 16)]
        string WorkspaceLastChangedBy { get; }

        /// <summary>
        /// SupportByVersion Office 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff864917.aspx </remarks>
        [SupportByVersion("Office", 11, 12, 14, 15, 16)]
        object LastSyncTime { get; }

        /// <summary>
        /// SupportByVersion Office 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff862150.aspx </remarks>
        [SupportByVersion("Office", 11, 12, 14, 15, 16)]
        NetOffice.OfficeApi.Enums.MsoSyncErrorType ErrorType { get; }

        /// <summary>
        /// SupportByVersion Office 11, 12, 14, 15, 16
        /// Get
        /// Unknown COM Proxy
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff860559.aspx </remarks>
        [SupportByVersion("Office", 11, 12, 14, 15, 16), ProxyResult]
        object Parent { get; }

        #endregion

        #region Methods

        /// <summary>
        /// SupportByVersion Office 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff863651.aspx </remarks>
        [SupportByVersion("Office", 11, 12, 14, 15, 16)]
        void GetUpdate();

        /// <summary>
        /// SupportByVersion Office 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff860754.aspx </remarks>
        [SupportByVersion("Office", 11, 12, 14, 15, 16)]
        void PutUpdate();

        /// <summary>
        /// SupportByVersion Office 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff860783.aspx </remarks>
        /// <param name="syncVersionType">NetOffice.OfficeApi.Enums.MsoSyncVersionType syncVersionType</param>
        [SupportByVersion("Office", 11, 12, 14, 15, 16)]
        void OpenVersion(NetOffice.OfficeApi.Enums.MsoSyncVersionType syncVersionType);

        /// <summary>
        /// SupportByVersion Office 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff864675.aspx </remarks>
        /// <param name="syncConflictResolution">NetOffice.OfficeApi.Enums.MsoSyncConflictResolutionType syncConflictResolution</param>
        [SupportByVersion("Office", 11, 12, 14, 15, 16)]
        void ResolveConflict(NetOffice.OfficeApi.Enums.MsoSyncConflictResolutionType syncConflictResolution);

        /// <summary>
        /// SupportByVersion Office 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff861422.aspx </remarks>
        [SupportByVersion("Office", 11, 12, 14, 15, 16)]
        void Unsuspend();

        #endregion
    }
}
