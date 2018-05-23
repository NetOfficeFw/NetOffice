using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.OfficeApi
{
    /// <summary>
    /// DispatchInterface LanguageSettings 
    /// SupportByVersion Office, 9,10,11,12,14,15,16
    /// </summary>
    /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff863125.aspx </remarks>
    [SupportByVersion("Office", 9, 10, 11, 12, 14, 15, 16)]
    [EntityType(EntityType.IsDispatchInterface)]
    public interface LanguageSettings : _IMsoDispObj
    {
        #region Properties

        /// <summary>
        /// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff863438.aspx </remarks>
        /// <param name="id">NetOffice.OfficeApi.Enums.MsoAppLanguageID id</param>
        [SupportByVersion("Office", 9, 10, 11, 12, 14, 15, 16)]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        Int32 get_LanguageID(NetOffice.OfficeApi.Enums.MsoAppLanguageID id);

        /// <summary>
        /// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
        /// Alias for get_LanguageID
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff863438.aspx </remarks>
        /// <param name="id">NetOffice.OfficeApi.Enums.MsoAppLanguageID id</param>
        [SupportByVersion("Office", 9, 10, 11, 12, 14, 15, 16), Redirect("get_LanguageID")]
        Int32 LanguageID(NetOffice.OfficeApi.Enums.MsoAppLanguageID id);

        /// <summary>
        /// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff861143.aspx </remarks>
        /// <param name="lid">NetOffice.OfficeApi.Enums.MsoLanguageID lid</param>
        [SupportByVersion("Office", 9, 10, 11, 12, 14, 15, 16)]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        bool get_LanguagePreferredForEditing(NetOffice.OfficeApi.Enums.MsoLanguageID lid);

        /// <summary>
        /// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
        /// Alias for get_LanguagePreferredForEditing
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff861143.aspx </remarks>
        /// <param name="lid">NetOffice.OfficeApi.Enums.MsoLanguageID lid</param>
        [SupportByVersion("Office", 9, 10, 11, 12, 14, 15, 16), Redirect("get_LanguagePreferredForEditing")]
        bool LanguagePreferredForEditing(NetOffice.OfficeApi.Enums.MsoLanguageID lid);

        /// <summary>
        /// SupportByVersion Office 10, 11, 12, 14, 15, 16
        /// Get
        /// Unknown COM Proxy
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff862054.aspx </remarks>
        [SupportByVersion("Office", 10, 11, 12, 14, 15, 16), ProxyResult]
        object Parent { get; }

        #endregion
    }
}
