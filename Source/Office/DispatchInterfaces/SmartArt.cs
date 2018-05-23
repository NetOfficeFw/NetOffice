using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.OfficeApi
{
    /// <summary>
    /// DispatchInterface SmartArt 
    /// SupportByVersion Office, 14,15,16
    /// </summary>
    /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff860804.aspx </remarks>
    [SupportByVersion("Office", 14, 15, 16)]
    [EntityType(EntityType.IsDispatchInterface)]
    public interface SmartArt : _IMsoDispObj
    {
        #region Properties

        /// <summary>
        /// SupportByVersion Office 14, 15, 16
        /// Get
        /// Unknown COM Proxy
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff864691.aspx </remarks>
        [SupportByVersion("Office", 14, 15, 16), ProxyResult]
        object Parent { get; }

        /// <summary>
        /// SupportByVersion Office 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff862828.aspx </remarks>
        [SupportByVersion("Office", 14, 15, 16)]
        NetOffice.OfficeApi.SmartArtNodes AllNodes { get; }

        /// <summary>
        /// SupportByVersion Office 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff860244.aspx </remarks>
        [SupportByVersion("Office", 14, 15, 16)]
        NetOffice.OfficeApi.SmartArtNodes Nodes { get; }

        /// <summary>
        /// SupportByVersion Office 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff861866.aspx </remarks>
        [SupportByVersion("Office", 14, 15, 16)]
        NetOffice.OfficeApi.SmartArtLayout Layout { get; set; }

        /// <summary>
        /// SupportByVersion Office 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff862785.aspx </remarks>
        [SupportByVersion("Office", 14, 15, 16)]
        NetOffice.OfficeApi.SmartArtQuickStyle QuickStyle { get; set; }

        /// <summary>
        /// SupportByVersion Office 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff862120.aspx </remarks>
        [SupportByVersion("Office", 14, 15, 16)]
        NetOffice.OfficeApi.SmartArtColor Color { get; set; }

        /// <summary>
        /// SupportByVersion Office 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff865245.aspx </remarks>
        [SupportByVersion("Office", 14, 15, 16)]
        NetOffice.OfficeApi.Enums.MsoTriState Reverse { get; set; }

        #endregion

        #region Methods

        /// <summary>
        /// SupportByVersion Office 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff864968.aspx </remarks>
        [SupportByVersion("Office", 14, 15, 16)]
        void Reset();

        #endregion
    }
}
