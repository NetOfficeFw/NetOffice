using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.OfficeApi
{
    /// <summary>
    /// DispatchInterface Crop 
    /// SupportByVersion Office, 14,15,16
    /// </summary>
    /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff860761.aspx </remarks>
    [SupportByVersion("Office", 14, 15, 16)]
    [EntityType(EntityType.IsDispatchInterface)]
    public interface Crop : _IMsoDispObj
    {
        #region Properties

        /// <summary>
        /// SupportByVersion Office 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff862450.aspx </remarks>
        [SupportByVersion("Office", 14, 15, 16)]
        Single PictureOffsetX { get; set; }

        /// <summary>
        /// SupportByVersion Office 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff864637.aspx </remarks>
        [SupportByVersion("Office", 14, 15, 16)]
        Single PictureOffsetY { get; set; }

        /// <summary>
        /// SupportByVersion Office 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff860544.aspx </remarks>
        [SupportByVersion("Office", 14, 15, 16)]
        Single PictureWidth { get; set; }

        /// <summary>
        /// SupportByVersion Office 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff860512.aspx </remarks>
        [SupportByVersion("Office", 14, 15, 16)]
        Single PictureHeight { get; set; }

        /// <summary>
        /// SupportByVersion Office 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff861232.aspx </remarks>
        [SupportByVersion("Office", 14, 15, 16)]
        Single ShapeLeft { get; set; }

        /// <summary>
        /// SupportByVersion Office 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff861517.aspx </remarks>
        [SupportByVersion("Office", 14, 15, 16)]
        Single ShapeTop { get; set; }

        /// <summary>
        /// SupportByVersion Office 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff861716.aspx </remarks>
        [SupportByVersion("Office", 14, 15, 16)]
        Single ShapeWidth { get; set; }

        /// <summary>
        /// SupportByVersion Office 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff864643.aspx </remarks>
        [SupportByVersion("Office", 14, 15, 16)]
        Single ShapeHeight { get; set; }

        #endregion
    }
}
