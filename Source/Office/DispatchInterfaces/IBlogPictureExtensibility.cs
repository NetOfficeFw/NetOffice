using System.Reflection;
using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.OfficeApi
{
    /// <summary>
    /// DispatchInterface IBlogPictureExtensibility 
    /// SupportByVersion Office, 12,14,15,16
    /// </summary>
    /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff860265.aspx </remarks>
    [SupportByVersion("Office", 12, 14, 15, 16)]
    [EntityType(EntityType.IsDispatchInterface)]
	[TypeId("000C03C5-0000-0000-C000-000000000046")]
    public interface IBlogPictureExtensibility : ICOMObject
    {
        #region Methods

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff860839.aspx </remarks>
        /// <param name="blogPictureProvider">string blogPictureProvider</param>
        /// <param name="friendlyName">string friendlyName</param>
        [SupportByVersion("Office", 12, 14, 15, 16)]
        void BlogPictureProviderProperties(out string blogPictureProvider, out string friendlyName);

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff862798.aspx </remarks>
        /// <param name="account">string account</param>
        /// <param name="blogProvider">string blogProvider</param>
        /// <param name="parentWindow">Int32 parentWindow</param>
        /// <param name="document">object document</param>
        [SupportByVersion("Office", 12, 14, 15, 16)]
        void CreatePictureAccount(string account, string blogProvider, Int32 parentWindow, object document);

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff864012.aspx </remarks>
        /// <param name="account">string account</param>
        /// <param name="parentWindow">Int32 parentWindow</param>
        /// <param name="document">object document</param>
        /// <param name="image">object image</param>
        /// <param name="pictureURI">string pictureURI</param>
        /// <param name="imageType">Int32 imageType</param>
        [SupportByVersion("Office", 12, 14, 15, 16)]
        void PublishPicture(string account, Int32 parentWindow, object document, object image, out string pictureURI, Int32 imageType);

        #endregion
    }
}
