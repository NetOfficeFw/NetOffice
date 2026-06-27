using System;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;

namespace NetOffice.OfficeApi.Native
{
    /// <summary>
    /// Represents an object that provides the ability to manipulate blog images.
    /// </summary>
    /// <remarks> https://msdn.microsoft.com/en-us/library/microsoft.office.core.iblogpictureextensibility.aspx </remarks>
    [ComImport, Guid("000C03C5-0000-0000-C000-000000000046"), TypeLibType(4288)]
    public interface IBlogPictureExtensibility
    {
        /// <summary>
        /// Enables picture providers to offer themselves as an upload location for blog pictures.
        /// </summary>
        /// <param name="BlogPictureProvider">The ID of the picture provider.</param>
        /// <param name="FriendlyName">The friendly name of the picture provider.</param>
        /// <remarks> https://msdn.microsoft.com/en-us/library/microsoft.office.core.iblogpictureextensibility.blogpictureproviderproperties.aspx </remarks>
        [DispId(1)]
        [MethodImpl(4096)]
        void BlogPictureProviderProperties([MarshalAs(19)] out string BlogPictureProvider, [MarshalAs(19)] out string FriendlyName);

        /// <summary>
        /// Allows a picture provider to display the user interface needed to guide the user through setting up a picture account.
        /// </summary>
        /// <remarks> https://msdn.microsoft.com/en-us/library/microsoft.office.core.iblogpictureextensibility.createpictureaccount.aspx </remarks>
        /// <param name="Account">Represents the GUID of the account registry key. Blog account settings are stored in the registry at \\HKCU\Software\Microsoft\Office\Common\Blog\Account.</param>
        /// <param name="BlogProvider">The ID of the provider.</param>
        /// <param name="ParentWindow">Contains the HWND for the window Microsoft Office Word is calling from.</param>
        /// <param name="Document">The current document.</param>
        [DispId(2)]
        [MethodImpl(4096)]
        void CreatePictureAccount([MarshalAs(19)][In] string Account, [MarshalAs(19)][In] string BlogProvider, [In] int ParentWindow, [MarshalAs(26)][In] object Document);

        /// <summary>
        /// Used to post a picture object to its final destination in a blog.
        /// </summary>
        /// <remarks> https://msdn.microsoft.com/en-us/library/microsoft.office.core.iblogpictureextensibility.publishpicture.aspx </remarks>
        /// <param name="Account">Represents the GUID of the account registry key. Blog account settings are stored in the registry at \\HKCU\Software\Microsoft\Office\Common\Blog\Account.</param>
        /// <param name="ParentWindow">Contains the HWND for the window Microsoft Office Word is calling from.</param>
        /// <param name="Document">The current document.</param>
        /// <param name="Image">Represents the name of the image file.</param>
        /// <param name="PictureURI">The URI of the picture.</param>
        /// <param name="ImageType">no description available</param>
        [DispId(3)]
        [MethodImpl(4096)]
        void PublishPicture([MarshalAs(19)][In] string Account, [In] int ParentWindow, [MarshalAs(26)][In] object Document, [MarshalAs(25)][In] object Image, [MarshalAs(19)] out string PictureURI, [In] int ImageType);
    }
}
