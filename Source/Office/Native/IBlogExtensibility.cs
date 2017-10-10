using System;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using NetOffice.OfficeApi.Enums;

namespace NetOffice.OfficeApi.Native
{
    /// <summary>
    /// An object that provides the ability to manipulate blog entries.
    /// </summary>
    /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff863146.aspx </remarks>
    [ComImport, Guid("000C03C4-0000-0000-C000-000000000046"), TypeLibType(4288)]
    public interface IBlogExtensibility
    {
        /// <summary>
        /// Contains information about the provider.
        /// </summary>
        /// <remarks> https://msdn.microsoft.com/VBA/Office-Shared-VBA/articles/iblogextensibility-blogproviderproperties-method-office </remarks>
        /// <param name="BlogProvider">The name of the blog provider.</param>
        /// <param name="FriendlyName">Represents the name displayed in the user interface.</param>
        /// <param name="CategorySupport">Represents how many categories are supported by the provider.</param>
        /// <param name="Padding">Specifies whether table padding is recognized.</param>
        [DispId(1)]
        [MethodImpl(4096)]
        void BlogProviderProperties([MarshalAs(19)] out string BlogProvider, [MarshalAs(19)] out string FriendlyName, out MsoBlogCategorySupport CategorySupport, out bool Padding);

        /// <summary>
        /// Called from the Choose Account dialog when the provider's name is chosen in the Blog Host dropdown or when the user requests to change a provider's account in the Blog Accounts dialog box.
        /// </summary>
        /// <remarks> https://msdn.microsoft.com/VBA/Office-Shared-VBA/articles/iblogextensibility-setupblogaccount-method-office </remarks> 
        /// <param name="Account">Represents the GUID of the account registry key. Blog account settings are stored in the registry at \HKCU\Software\Microsoft\Office\Common\Blog\Account.</param>
        /// <param name="ParentWindow">Contains the HWND for the window Microsoft Word is calling from.</param>
        /// <param name="Document">The current document.</param>
        /// <param name="NewAccount">Indicates whether this is a new account.</param>
        /// <param name="ShowPictureUI">Indicates whether Microsoft Word's picture user interface needs to be displayed.</param>
        [DispId(2)]
        [MethodImpl(4096)]
        void SetupBlogAccount([MarshalAs(19)] [In] string Account, [In] int ParentWindow, [MarshalAs(26)] [In] object Document, [In] bool NewAccount, out bool ShowPictureUI);

        /// <summary>
        /// Returns the list and details of user blogs associated with the specified account.
        /// </summary>
        /// <remarks> https://msdn.microsoft.com/VBA/Office-Shared-VBA/articles/iblogextensibility-getuserblogs-method-office </remarks>
        /// <param name="Account">Represents the GUID of the account registry key. Blog account settings are stored in the registry at \HKCU\Software\Microsoft\Office\Common\Blog\Account.</param>
        /// <param name="ParentWindow">Contains the HWND for the window Microsoft Word is calling from.</param>
        /// <param name="Document">The current document.</param>
        /// <param name="BlogNames">Contains all blog names under the current account.</param>
        /// <param name="BlogIDs">Contains all blog IDs under the current account.</param>
        /// <param name="BlogURLs">Contains all blog URLs under the current account.</param>
        [DispId(3)]
        [MethodImpl(4096)]
        void GetUserBlogs([MarshalAs(19)] [In] string Account, [In] int ParentWindow, [MarshalAs(26)] [In] object Document, [MarshalAs(29, SafeArraySubType = VarEnum.VT_BSTR)] out Array BlogNames, [MarshalAs(29, SafeArraySubType = VarEnum.VT_BSTR)] out Array BlogIDs, [MarshalAs(29, SafeArraySubType = VarEnum.VT_BSTR)] out Array BlogURLs);

        /// <summary>
        /// Returns the list of the user's last fifteen blog posts that Microsoft Word then displays in the Open Existing Post dialog. This method does not actually return the blog post contents.
        /// </summary>
        /// <remarks> https://msdn.microsoft.com/VBA/Office-Shared-VBA/articles/iblogextensibility-getrecentposts-method-office </remarks>
        /// <param name="Account">Represents the GUID of the account registry key. Blog account settings are stored in the registry at \HKCU\Software\Microsoft\Office\Common\Blog\Account.</param>
        /// <param name="ParentWindow">Contains the HWND for the window Microsoft Word is calling from.</param>
        /// <param name="Document">The current document.</param>
        /// <param name="PostTitles">Contains the titles of the last fifteen posts.</param>
        /// <param name="PostDates">Contains the dates of the last fifteen posts.</param>
        /// <param name="PostIDs">Contains the IDs of the last fifteen posts.</param>
        [DispId(4)]
        [MethodImpl(4096)]
        void GetRecentPosts([MarshalAs(19)] [In] string Account, [In] int ParentWindow, [MarshalAs(26)] [In] object Document, [MarshalAs(29, SafeArraySubType = VarEnum.VT_BSTR)] out Array PostTitles, [MarshalAs(29, SafeArraySubType = VarEnum.VT_BSTR)] out Array PostDates, [MarshalAs(29, SafeArraySubType = VarEnum.VT_BSTR)] out Array PostIDs);

        /// <summary>
        /// Opens the blog specified by the blog ID. It is called by the Open Existing Post dialog based on the item selected by the user.
        /// </summary>
        /// <remarks> https://msdn.microsoft.com/VBA/Office-Shared-VBA/articles/iblogextensibility-open-method-office </remarks>
        /// <param name="Account">Represents the GUID of the account registry key. Blog account settings are stored in the registry at \HKCU\Software\Microsoft\Office\Common\Blog\Account.</param>
        /// <param name="PostID">The ID of the post.</param>
        /// <param name="ParentWindow">Contains the HWND for the window Microsoft Word is calling from.</param>
        /// <param name="xHTML">Represents the xHTML of the current document.</param>
        /// <param name="Title">The title of the post.</param>
        /// <param name="DatePosted">The date the entry was posted.</param>
        /// <param name="Categories">A list of categories supported by the provider.</param>
        [DispId(5)]
        [MethodImpl(4096)]
        void Open([MarshalAs(19)] [In] string Account, [MarshalAs(19)] [In] string PostID, [In] int ParentWindow, [MarshalAs(19)] out string xHTML, [MarshalAs(19)] out string Title, [MarshalAs(19)] out string DatePosted, [MarshalAs(29, SafeArraySubType = VarEnum.VT_BSTR)] out Array Categories);

        /// <summary>
        /// Hands off the current post so it can be published by the provider.
        /// </summary>
        /// <remarks> https://msdn.microsoft.com/VBA/Office-Shared-VBA/articles/iblogextensibility-publishpost-method-office </remarks>
        /// <param name="Account">Represents the GUID of the account registry key. Blog account settings are stored in the registry at \HKCU\Software\Microsoft\Office\Common\Blog\Account.</param>
        /// <param name="ParentWindow">Contains the HWND for the window Microsoft Word is calling from.</param>
        /// <param name="Document">The current document.</param>
        /// <param name="xHTML">Represents the xHTML of the current document.</param>
        /// <param name="Title">The title of the post.</param>
        /// <param name="DateTime">The date the entry was posted.</param>
        /// <param name="Categories">A list of categories supported by the provider.</param>
        /// <param name="Draft">Specifies whether this is a draft version of the post.</param>
        /// <param name="PostID">The ID of the original post if this post has been republished.</param>
        /// <param name="PublishMessage">Specifies what is displayed in the publish bar.</param>
        [DispId(6)]
        [MethodImpl(4096)]
        void PublishPost([MarshalAs(19)] [In] string Account, [In] int ParentWindow, [MarshalAs(26)] [In] object Document, [MarshalAs(19)] [In] string xHTML, [MarshalAs(19)] [In] string Title, [MarshalAs(19)] [In] string DateTime, [MarshalAs(29, SafeArraySubType = VarEnum.VT_BSTR)] [In] Array Categories, [In] bool Draft, [MarshalAs(19)] out string PostID, [MarshalAs(19)] out string PublishMessage);

        /// <summary>
        /// Hands off the current post so it can be republished by the provider.
        /// </summary>
        /// <remarks> https://msdn.microsoft.com/VBA/Office-Shared-VBA/articles/iblogextensibility-republishpost-method-office </remarks>
        /// <param name="Account">Represents the GUID of the account registry key. Blog account settings are stored in the registry at \HKCU\Software\Microsoft\Office\Common\Blog\Account.</param>
        /// <param name="ParentWindow">Contains the HWND for the window Microsoft Word is calling from.</param>
        /// <param name="Document">The current document.</param>
        /// <param name="PostID">The ID of the original post.</param>
        /// <param name="xHTML">Represents the xHTML of the current document.</param>
        /// <param name="Title">The title of the post.</param>
        /// <param name="DateTime">The date the entry was posted.</param>
        /// <param name="Categories">A list of categories supported by the provider.</param>
        /// <param name="Draft">Specifies whether this is a draft version of the post.</param>
        /// <param name="PublishMessage">Specifies what is displayed in the publish bar.</param>
        [DispId(7)]
        [MethodImpl(4096)]
        void RepublishPost([MarshalAs(19)] [In] string Account, [In] int ParentWindow, [MarshalAs(26)] [In] object Document, [MarshalAs(19)] [In] string PostID, [MarshalAs(19)] [In] string xHTML, [MarshalAs(19)] [In] string Title, [MarshalAs(19)] [In] string DateTime, [MarshalAs(29, SafeArraySubType = VarEnum.VT_BSTR)] [In] Array Categories, [In] bool Draft, [MarshalAs(19)] out string PublishMessage);

        /// <summary>
        /// This method returns the list of blog categories for an account so Microsoft Word can populate the categories dropdown list.
        /// </summary>
        /// <remarks> https://msdn.microsoft.com/VBA/Office-Shared-VBA/articles/iblogextensibility-getcategories-method-office </remarks>
        /// <param name="Account">Represents the GUID of the account registry key.</param>
        /// <param name="ParentWindow">Represents the HWND of the host window.</param>
        /// <param name="Document">The current document.</param>
        /// <param name="Categories">A list of categories supported by the provider.</param>
        [DispId(8)]
        [MethodImpl(4096)]
        void GetCategories([MarshalAs(19)] [In] string Account, [In] int ParentWindow, [MarshalAs(26)] [In] object Document, [MarshalAs(29, SafeArraySubType = VarEnum.VT_BSTR)] out Array Categories);
    }
}