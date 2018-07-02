using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.OfficeApi
{
    /// <summary>
    /// _CustomTaskPane AfterDelete Event Handler
    /// AfterDelete is a custom event from NetOffice
    /// </summary>
    /// <param name="sender">deleted pane</param>
    public delegate void _CustomTaskPaneDeleteHandler(_CustomTaskPane sender);

    /// <summary>
    /// DispatchInterface _CustomTaskPane
    /// SupportByVersion Office, 12,14,15,16
    /// </summary>
    [SupportByVersion("Office", 12, 14, 15, 16)]
    [EntityType(EntityType.IsDispatchInterface), BaseType]
	[TypeId("000C033B-0000-0000-C000-000000000046")]
    [CoClassSource(typeof(NetOffice.OfficeApi.CustomTaskPane))]
    public interface _CustomTaskPane : ICOMObject
    {
        #region Events

        /// <summary>
        /// Occurs after Delete for the proxy has been called.
        /// This is a custom event from NetOffice.
        /// </summary>
        /// <remarks>The event occurs for the proxy instance only.</remarks>
        [CustomEvent]
        event _CustomTaskPaneDeleteHandler AfterDelete;

        #endregion

        #region Properties

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff861137.aspx </remarks>
        [SupportByVersion("Office", 12, 14, 15, 16)]
        string Title { get; }

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// Get
        /// Unknown COM Proxy
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff862545.aspx </remarks>
        [SupportByVersion("Office", 12, 14, 15, 16), ProxyResult]
        object Application { get; }

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// Get
        /// Unknown COM Proxy
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff862803.aspx </remarks>
        [SupportByVersion("Office", 12, 14, 15, 16), ProxyResult]
        object Window { get; }

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff865256.aspx </remarks>
        [SupportByVersion("Office", 12, 14, 15, 16)]
        bool Visible { get; set; }

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// Get
        /// Native COM Proxy
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff861783.aspx </remarks>
        [SupportByVersion("Office", 12, 14, 15, 16), NativeResult]
        object ContentControl { get; }

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff860235.aspx </remarks>
        [SupportByVersion("Office", 12, 14, 15, 16)]
        Int32 Height { get; set; }

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff865362.aspx </remarks>
        [SupportByVersion("Office", 12, 14, 15, 16)]
        Int32 Width { get; set; }

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff861841.aspx </remarks>
        [SupportByVersion("Office", 12, 14, 15, 16)]
        NetOffice.OfficeApi.Enums.MsoCTPDockPosition DockPosition { get; set; }


        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff861088.aspx </remarks>
        [SupportByVersion("Office", 12, 14, 15, 16)]
        NetOffice.OfficeApi.Enums.MsoCTPDockPositionRestrict DockPositionRestrict { get; set; }

        #endregion

        #region Methods

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff862399.aspx </remarks>
        [SupportByVersion("Office", 12, 14, 15, 16)]
        void Delete();

        #endregion
    }
}
