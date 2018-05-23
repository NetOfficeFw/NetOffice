using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.OfficeApi
{
    /// <summary>
    /// DispatchInterface PictureEffect 
    /// SupportByVersion Office, 14,15,16
    /// </summary>
    /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff863789.aspx </remarks>
    [SupportByVersion("Office", 14, 15, 16)]
    [EntityType(EntityType.IsDispatchInterface)]
    public interface PictureEffect : _IMsoDispObj
    {
        #region Properties

        /// <summary>
        /// SupportByVersion Office 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff863360.aspx </remarks>
        [SupportByVersion("Office", 14, 15, 16)]
        NetOffice.OfficeApi.Enums.MsoPictureEffectType Type { get; }

        /// <summary>
        /// SupportByVersion Office 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff860898.aspx </remarks>
        [SupportByVersion("Office", 14, 15, 16)]
        Int32 Position { get; set; }

        /// <summary>
        /// SupportByVersion Office 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff863402.aspx </remarks>
        [SupportByVersion("Office", 14, 15, 16)]
        NetOffice.OfficeApi.EffectParameters EffectParameters { get; }

        /// <summary>
        /// SupportByVersion Office 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff864650.aspx </remarks>
        [SupportByVersion("Office", 14, 15, 16)]
        NetOffice.OfficeApi.Enums.MsoTriState Visible { get; set; }

        #endregion

        #region Methods

        /// <summary>
        /// SupportByVersion Office 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff864645.aspx </remarks>
        [SupportByVersion("Office", 14, 15, 16)]
        void Delete();

        #endregion
    }
}
