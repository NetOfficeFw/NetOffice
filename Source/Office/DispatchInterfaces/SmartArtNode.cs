using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.OfficeApi
{
    /// <summary>
    /// DispatchInterface SmartArtNode 
    /// SupportByVersion Office, 14,15,16
    /// </summary>
    /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff861178.aspx </remarks>
    [SupportByVersion("Office", 14, 15, 16)]
    [EntityType(EntityType.IsDispatchInterface)]
    public interface SmartArtNode : _IMsoDispObj
    {
        #region Properties

        /// <summary>
        /// SupportByVersion Office 14, 15, 16
        /// Get
        /// Unknown COM Proxy
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff863308.aspx </remarks>
        [SupportByVersion("Office", 14, 15, 16), ProxyResult]
        object Parent { get; }

        /// <summary>
        /// SupportByVersion Office 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff860568.aspx </remarks>
        [SupportByVersion("Office", 14, 15, 16)]
        NetOffice.OfficeApi.Enums.MsoOrgChartLayoutType OrgChartLayout { get; set; }

        /// <summary>
        /// SupportByVersion Office 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff864604.aspx </remarks>
        [SupportByVersion("Office", 14, 15, 16)]
        NetOffice.OfficeApi.ShapeRange Shapes { get; }

        /// <summary>
        /// SupportByVersion Office 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff861779.aspx </remarks>
        [SupportByVersion("Office", 14, 15, 16)]
        NetOffice.OfficeApi.TextFrame2 TextFrame2 { get; }

        /// <summary>
        /// SupportByVersion Office 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff862082.aspx </remarks>
        [SupportByVersion("Office", 14, 15, 16)]
        Int32 Level { get; }

        /// <summary>
        /// SupportByVersion Office 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff860275.aspx </remarks>
        [SupportByVersion("Office", 14, 15, 16)]
        NetOffice.OfficeApi.Enums.MsoTriState Hidden { get; }

        /// <summary>
        /// SupportByVersion Office 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff865275.aspx </remarks>
        [SupportByVersion("Office", 14, 15, 16)]
        NetOffice.OfficeApi.SmartArtNodes Nodes { get; }

        /// <summary>
        /// SupportByVersion Office 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff862047.aspx </remarks>
        [SupportByVersion("Office", 14, 15, 16)]
        NetOffice.OfficeApi.SmartArtNode ParentNode { get; }

        /// <summary>
        /// SupportByVersion Office 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff861873.aspx </remarks>
        [SupportByVersion("Office", 14, 15, 16)]
        NetOffice.OfficeApi.Enums.MsoSmartArtNodeType Type { get; }

        #endregion

        #region Methods

        /// <summary>
        /// SupportByVersion Office 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff865366.aspx </remarks>
        /// <param name="position">optional NetOffice.OfficeApi.Enums.MsoSmartArtNodePosition Position = 1</param>
        /// <param name="type">optional NetOffice.OfficeApi.Enums.MsoSmartArtNodeType Type = 1</param>
        [SupportByVersion("Office", 14, 15, 16)]
        NetOffice.OfficeApi.SmartArtNode AddNode(object position, object type);

        /// <summary>
        /// SupportByVersion Office 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff865366.aspx </remarks>
        [CustomMethod]
        [SupportByVersion("Office", 14, 15, 16)]
        NetOffice.OfficeApi.SmartArtNode AddNode();

        /// <summary>
        /// SupportByVersion Office 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff865366.aspx </remarks>
        /// <param name="position">optional NetOffice.OfficeApi.Enums.MsoSmartArtNodePosition Position = 1</param>
        [CustomMethod]
        [SupportByVersion("Office", 14, 15, 16)]
        NetOffice.OfficeApi.SmartArtNode AddNode(object position);

        /// <summary>
        /// SupportByVersion Office 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff863109.aspx </remarks>
        [SupportByVersion("Office", 14, 15, 16)]
        void Delete();

        /// <summary>
        /// SupportByVersion Office 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff862804.aspx </remarks>
        [SupportByVersion("Office", 14, 15, 16)]
        void Promote();

        /// <summary>
        /// SupportByVersion Office 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff860258.aspx </remarks>
        [SupportByVersion("Office", 14, 15, 16)]
        void Demote();

        /// <summary>
        /// SupportByVersion Office 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff864694.aspx </remarks>
        [SupportByVersion("Office", 14, 15, 16)]
        void Larger();

        /// <summary>
        /// SupportByVersion Office 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff863061.aspx </remarks>
        [SupportByVersion("Office", 14, 15, 16)]
        void Smaller();

        /// <summary>
        /// SupportByVersion Office 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff863035.aspx </remarks>
        [SupportByVersion("Office", 14, 15, 16)]
        void ReorderUp();

        /// <summary>
        /// SupportByVersion Office 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff860343.aspx </remarks>
        [SupportByVersion("Office", 14, 15, 16)]
        void ReorderDown();

        #endregion
    }
}
