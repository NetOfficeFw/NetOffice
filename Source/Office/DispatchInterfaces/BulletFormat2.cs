using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.OfficeApi
{
    /// <summary>
    /// DispatchInterface BulletFormat2 
    /// SupportByVersion Office, 12,14,15,16
    /// </summary>
    /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff863746.aspx </remarks>
    [SupportByVersion("Office", 12, 14, 15, 16)]
    [EntityType(EntityType.IsDispatchInterface)]
    public interface BulletFormat2 : _IMsoDispObj
    {
        #region Properties

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// Get
        /// Unknown COM Proxy
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff861186.aspx </remarks>
        [SupportByVersion("Office", 12, 14, 15, 16), ProxyResult]
        object Parent { get; }

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff860608.aspx </remarks>
        [SupportByVersion("Office", 12, 14, 15, 16)]
        Int32 Character { get; set; }

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff860536.aspx </remarks>
        [SupportByVersion("Office", 12, 14, 15, 16)]
        NetOffice.OfficeApi.Font2 Font { get; }

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff864635.aspx </remarks>
        [SupportByVersion("Office", 12, 14, 15, 16)]
        Int32 Number { get; }

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff863117.aspx </remarks>
        [SupportByVersion("Office", 12, 14, 15, 16)]
        Single RelativeSize { get; set; }

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff860774.aspx </remarks>
        [SupportByVersion("Office", 12, 14, 15, 16)]
        Int32 StartValue { get; set; }

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff860844.aspx </remarks>
        [SupportByVersion("Office", 12, 14, 15, 16)]
        NetOffice.OfficeApi.Enums.MsoNumberedBulletStyle Style { get; set; }

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff863284.aspx </remarks>
        [SupportByVersion("Office", 12, 14, 15, 16)]
        NetOffice.OfficeApi.Enums.MsoBulletType Type { get; set; }

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff864181.aspx </remarks>
        [SupportByVersion("Office", 12, 14, 15, 16)]
        NetOffice.OfficeApi.Enums.MsoTriState UseTextColor { get; set; }

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff861075.aspx </remarks>
        [SupportByVersion("Office", 12, 14, 15, 16)]
        NetOffice.OfficeApi.Enums.MsoTriState UseTextFont { get; set; }

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff861165.aspx </remarks>
        [SupportByVersion("Office", 12, 14, 15, 16)]
        NetOffice.OfficeApi.Enums.MsoTriState Visible { get; set; }

        #endregion

        #region Methods

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff860556.aspx </remarks>
        /// <param name="fileName">string fileName</param>
        [SupportByVersion("Office", 12, 14, 15, 16)]
        void Picture(string fileName);

        #endregion
    }
}
