using System.Reflection;
using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.OfficeApi
{
    /// <summary>
    /// DispatchInterface IBlogExtensibility 
    /// SupportByVersion Office, 12,14,15,16
    /// </summary>
    /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff863146.aspx </remarks>
    [SupportByVersion("Office", 12, 14, 15, 16)]
    [EntityType(EntityType.IsDispatchInterface)]
	[TypeId("000C03C4-0000-0000-C000-000000000046")]
    public interface IBlogExtensibility : ICOMObject
    {
        #region Methods

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff862840.aspx </remarks>
        /// <param name="blogProvider">string blogProvider</param>
        /// <param name="friendlyName">string friendlyName</param>
        /// <param name="categorySupport">NetOffice.OfficeApi.Enums.MsoBlogCategorySupport categorySupport</param>
        /// <param name="padding">bool padding</param>
        [SupportByVersion("Office", 12, 14, 15, 16)]
        void BlogProviderProperties(out string blogProvider, out string friendlyName, out NetOffice.OfficeApi.Enums.MsoBlogCategorySupport categorySupport, out bool padding);

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff863154.aspx </remarks>
        /// <param name="account">string account</param>
        /// <param name="parentWindow">Int32 parentWindow</param>
        /// <param name="document">object document</param>
        /// <param name="newAccount">bool newAccount</param>
        /// <param name="showPictureUI">bool showPictureUI</param>
        [SupportByVersion("Office", 12, 14, 15, 16)]
        void SetupBlogAccount(string account, Int32 parentWindow, object document, bool newAccount, out bool showPictureUI);

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff860220.aspx </remarks>
        /// <param name="account">string account</param>
        /// <param name="parentWindow">Int32 parentWindow</param>
        /// <param name="document">object document</param>
        /// <param name="blogNames">String[] blogNames</param>
        /// <param name="blogIDs">String[] blogIDs</param>
        /// <param name="blogURLs">String[] blogURLs</param>
        [SupportByVersion("Office", 12, 14, 15, 16)]
        void GetUserBlogs(string account, Int32 parentWindow, object document, out String[] blogNames, out String[] blogIDs, out String[] blogURLs);

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff861430.aspx </remarks>
        /// <param name="account">string account</param>
        /// <param name="parentWindow">Int32 parentWindow</param>
        /// <param name="document">object document</param>
        /// <param name="postTitles">String[] postTitles</param>
        /// <param name="postDates">String[] postDates</param>
        /// <param name="postIDs">String[] postIDs</param>
        [SupportByVersion("Office", 12, 14, 15, 16)]
        void GetRecentPosts(string account, Int32 parentWindow, object document, out String[] postTitles, out String[] postDates, out String[] postIDs);

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff861145.aspx </remarks>
        /// <param name="account">string account</param>
        /// <param name="postID">string postID</param>
        /// <param name="parentWindow">Int32 parentWindow</param>
        /// <param name="xHTML">string xHTML</param>
        /// <param name="title">string title</param>
        /// <param name="datePosted">string datePosted</param>
        /// <param name="categories">String[] categories</param>
        [SupportByVersion("Office", 12, 14, 15, 16)]
       void Open(string account, string postID, Int32 parentWindow, out string xHTML, out string title, out string datePosted, out String[] categories);

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff862458.aspx </remarks>
        /// <param name="account">string account</param>
        /// <param name="parentWindow">Int32 parentWindow</param>
        /// <param name="document">object document</param>
        /// <param name="xHTML">string xHTML</param>
        /// <param name="title">string title</param>
        /// <param name="dateTime">string dateTime</param>
        /// <param name="categories">String[] categories</param>
        /// <param name="draft">bool draft</param>
        /// <param name="postID">string postID</param>
        /// <param name="publishMessage">string publishMessage</param>
        [SupportByVersion("Office", 12, 14, 15, 16)]
        void PublishPost(string account, Int32 parentWindow, object document, string xHTML, string title, string dateTime, String[] categories, bool draft, out string postID, out string publishMessage);

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff860616.aspx </remarks>
        /// <param name="account">string account</param>
        /// <param name="parentWindow">Int32 parentWindow</param>
        /// <param name="document">object document</param>
        /// <param name="postID">string postID</param>
        /// <param name="xHTML">string xHTML</param>
        /// <param name="title">string title</param>
        /// <param name="dateTime">string dateTime</param>
        /// <param name="categories">String[] categories</param>
        /// <param name="draft">bool draft</param>
        /// <param name="publishMessage">string publishMessage</param>
        [SupportByVersion("Office", 12, 14, 15, 16)]
        void RepublishPost(string account, Int32 parentWindow, object document, string postID, string xHTML, string title, string dateTime, String[] categories, bool draft, out string publishMessage);

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff865355.aspx </remarks>
        /// <param name="account">string account</param>
        /// <param name="parentWindow">Int32 parentWindow</param>
        /// <param name="document">object document</param>
        /// <param name="categories">String[] categories</param>
        [SupportByVersion("Office", 12, 14, 15, 16)]
        void GetCategories(string account, Int32 parentWindow, object document, out String[] categories);

        #endregion
    }
}
