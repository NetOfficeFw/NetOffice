using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.OfficeApi
{
    /// <summary>
    /// DispatchInterface SharedWorkspace 
    /// SupportByVersion Office, 11,12,14,15,16
    /// </summary>
    /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff862502.aspx </remarks>
    [SupportByVersion("Office", 11, 12, 14, 15, 16)]
    [EntityType(EntityType.IsDispatchInterface)]
    public interface SharedWorkspace : _IMsoDispObj
    {
        #region Properties

        /// <summary>
        /// SupportByVersion Office 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff861084.aspx </remarks>
        [SupportByVersion("Office", 11, 12, 14, 15, 16)]
        string Name { get; set; }

        /// <summary>
        /// SupportByVersion Office 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff863506.aspx </remarks>
        [SupportByVersion("Office", 11, 12, 14, 15, 16)]
        NetOffice.OfficeApi.SharedWorkspaceMembers Members { get; }

        /// <summary>
        /// SupportByVersion Office 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff863392.aspx </remarks>
        [SupportByVersion("Office", 11, 12, 14, 15, 16)]
        NetOffice.OfficeApi.SharedWorkspaceTasks Tasks { get; }

        /// <summary>
        /// SupportByVersion Office 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff865183.aspx </remarks>
        [SupportByVersion("Office", 11, 12, 14, 15, 16)]
        NetOffice.OfficeApi.SharedWorkspaceFiles Files { get; }

        /// <summary>
        /// SupportByVersion Office 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff863702.aspx </remarks>
        [SupportByVersion("Office", 11, 12, 14, 15, 16)]
        NetOffice.OfficeApi.SharedWorkspaceFolders Folders { get; }

        /// <summary>
        /// SupportByVersion Office 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff862483.aspx </remarks>
        [SupportByVersion("Office", 11, 12, 14, 15, 16)]
        NetOffice.OfficeApi.SharedWorkspaceLinks Links { get; }

        /// <summary>
        /// SupportByVersion Office 11, 12, 14, 15, 16
        /// Get
        /// Unknown COM Proxy
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff861765.aspx </remarks>
        [SupportByVersion("Office", 11, 12, 14, 15, 16), ProxyResult]
        object Parent { get; }

        /// <summary>
        /// SupportByVersion Office 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff865214.aspx </remarks>
        [SupportByVersion("Office", 11, 12, 14, 15, 16)]
        string URL { get; }

        /// <summary>
        /// SupportByVersion Office 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff860257.aspx </remarks>
        [SupportByVersion("Office", 11, 12, 14, 15, 16)]
        bool Connected { get; }

        /// <summary>
        /// SupportByVersion Office 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff861389.aspx </remarks>
        [SupportByVersion("Office", 11, 12, 14, 15, 16)]
        object LastRefreshed { get; }

        /// <summary>
        /// SupportByVersion Office 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff860917.aspx </remarks>
        [SupportByVersion("Office", 11, 12, 14, 15, 16)]
        string SourceURL { get; set; }

        #endregion

        #region Methods

        /// <summary>
        /// SupportByVersion Office 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff862068.aspx </remarks>
        [SupportByVersion("Office", 11, 12, 14, 15, 16)]
        void Refresh();

        /// <summary>
        /// SupportByVersion Office 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff862182.aspx </remarks>
        /// <param name="uRL">optional object uRL</param>
        /// <param name="name">optional object name</param>
        [SupportByVersion("Office", 11, 12, 14, 15, 16)]
        void CreateNew(object uRL, object name);

        /// <summary>
        /// SupportByVersion Office 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff862182.aspx </remarks>
        [CustomMethod]
        [SupportByVersion("Office", 11, 12, 14, 15, 16)]
        void CreateNew();

        /// <summary>
        /// SupportByVersion Office 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff862182.aspx </remarks>
        /// <param name="uRL">optional object uRL</param>
        [CustomMethod]
        [SupportByVersion("Office", 11, 12, 14, 15, 16)]
        void CreateNew(object uRL);

        /// <summary>
        /// SupportByVersion Office 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff862550.aspx </remarks>
        [SupportByVersion("Office", 11, 12, 14, 15, 16)]
        void Delete();

        /// <summary>
        /// SupportByVersion Office 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff861519.aspx </remarks>
        [SupportByVersion("Office", 11, 12, 14, 15, 16)]
        void RemoveDocument();

        /// <summary>
        /// SupportByVersion Office 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff863540.aspx </remarks>
        [SupportByVersion("Office", 11, 12, 14, 15, 16)]
        void Disconnect();

        #endregion
    }
}
