using System;
using NetRuntimeSystem = System;
using System.Runtime.InteropServices;
using System.Runtime.CompilerServices;
using System.ComponentModel;
using System.Reflection;
using System.Collections.Generic;
using NetOffice;
namespace NetOffice.OfficeApi
{
	///<summary>
	/// DispatchInterface IBlogExtensibility 
	/// SupportByVersion Office, 12,14,15,16
	/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff863146.aspx
	///</summary>
	[SupportByVersionAttribute("Office", 12,14,15,16)]
	[EntityTypeAttribute(EntityType.IsDispatchInterface)]
	public class IBlogExtensibility : COMObject
	{
		#pragma warning disable
		#region Type Information

        private static Type _type;

		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        public static Type LateBindingApiWrapperType
        {
            get
            {
                if (null == _type)
                    _type = typeof(IBlogExtensibility);
                    
                return _type;
            }
        }
        
        #endregion
        
		#region Construction

		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
		public IBlogExtensibility(Core factory, ICOMObject parentObject, object comProxy) : base(factory, parentObject, comProxy)
		{
			
		}

        ///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public IBlogExtensibility(ICOMObject parentObject, object comProxy) : base(parentObject, comProxy)
		{
		}
		
		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public IBlogExtensibility(Core factory, ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(factory, parentObject, comProxy, comProxyType)
		{

		}

		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public IBlogExtensibility(ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(parentObject, comProxy, comProxyType)
		{
		}
		
		///<param name="replacedObject">object to replaced. replacedObject are not usable after this action</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public IBlogExtensibility(ICOMObject replacedObject) : base(replacedObject)
		{
		}
		
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public IBlogExtensibility() : base()
		{
		}
		
		/// <param name="progId">registered ProgID</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public IBlogExtensibility(string progId) : base(progId)
		{
		}
		
		#endregion
		
		#region Properties

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion Office 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff862840.aspx
		/// </summary>
		/// <param name="blogProvider">string BlogProvider</param>
		/// <param name="friendlyName">string FriendlyName</param>
		/// <param name="categorySupport">NetOffice.OfficeApi.Enums.MsoBlogCategorySupport CategorySupport</param>
		/// <param name="padding">bool Padding</param>
		[SupportByVersionAttribute("Office", 12,14,15,16)]
		public void BlogProviderProperties(out string blogProvider, out string friendlyName, out NetOffice.OfficeApi.Enums.MsoBlogCategorySupport categorySupport, out bool padding)
		{
			ParameterModifier[] modifiers = Invoker.CreateParamModifiers(true,true,true,true);
			blogProvider = string.Empty;
			friendlyName = string.Empty;
			categorySupport = 0;
			padding = false;
			object[] paramsArray = Invoker.ValidateParamsArray(blogProvider, friendlyName, categorySupport, padding);
			Invoker.Method(this, "BlogProviderProperties", paramsArray, modifiers);
			blogProvider = (string)paramsArray[0];
			friendlyName = (string)paramsArray[1];
			categorySupport = (NetOffice.OfficeApi.Enums.MsoBlogCategorySupport)paramsArray[2];
			padding = (bool)paramsArray[3];
		}

		/// <summary>
		/// SupportByVersion Office 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff863154.aspx
		/// </summary>
		/// <param name="account">string Account</param>
		/// <param name="parentWindow">Int32 ParentWindow</param>
		/// <param name="document">object Document</param>
		/// <param name="newAccount">bool NewAccount</param>
		/// <param name="showPictureUI">bool ShowPictureUI</param>
		[SupportByVersionAttribute("Office", 12,14,15,16)]
		public void SetupBlogAccount(string account, Int32 parentWindow, object document, bool newAccount, out bool showPictureUI)
		{
			ParameterModifier[] modifiers = Invoker.CreateParamModifiers(false,false,false,false,true);
			showPictureUI = false;
			object[] paramsArray = Invoker.ValidateParamsArray(account, parentWindow, document, newAccount, showPictureUI);
			Invoker.Method(this, "SetupBlogAccount", paramsArray, modifiers);
			showPictureUI = (bool)paramsArray[4];
		}

		/// <summary>
		/// SupportByVersion Office 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff860220.aspx
		/// </summary>
		/// <param name="account">string Account</param>
		/// <param name="parentWindow">Int32 ParentWindow</param>
		/// <param name="document">object Document</param>
		/// <param name="blogNames">String[] BlogNames</param>
		/// <param name="blogIDs">String[] BlogIDs</param>
		/// <param name="blogURLs">String[] BlogURLs</param>
		[SupportByVersionAttribute("Office", 12,14,15,16)]
		public void GetUserBlogs(string account, Int32 parentWindow, object document, out String[] blogNames, out String[] blogIDs, out String[] blogURLs)
		{
			ParameterModifier[] modifiers = Invoker.CreateParamModifiers(false,false,false,true,true,true);
			blogNames = null;
			blogIDs = null;
			blogURLs = null;
			object[] paramsArray = Invoker.ValidateParamsArray(account, parentWindow, document, (object)blogNames, (object)blogIDs, (object)blogURLs);
			Invoker.Method(this, "GetUserBlogs", paramsArray, modifiers);
			blogNames = (String[])paramsArray[3];
			blogIDs = (String[])paramsArray[4];
			blogURLs = (String[])paramsArray[5];
		}

		/// <summary>
		/// SupportByVersion Office 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff861430.aspx
		/// </summary>
		/// <param name="account">string Account</param>
		/// <param name="parentWindow">Int32 ParentWindow</param>
		/// <param name="document">object Document</param>
		/// <param name="postTitles">String[] PostTitles</param>
		/// <param name="postDates">String[] PostDates</param>
		/// <param name="postIDs">String[] PostIDs</param>
		[SupportByVersionAttribute("Office", 12,14,15,16)]
		public void GetRecentPosts(string account, Int32 parentWindow, object document, out String[] postTitles, out String[] postDates, out String[] postIDs)
		{
			ParameterModifier[] modifiers = Invoker.CreateParamModifiers(false,false,false,true,true,true);
			postTitles = null;
			postDates = null;
			postIDs = null;
			object[] paramsArray = Invoker.ValidateParamsArray(account, parentWindow, document, (object)postTitles, (object)postDates, (object)postIDs);
			Invoker.Method(this, "GetRecentPosts", paramsArray, modifiers);
			postTitles = (String[])paramsArray[3];
			postDates = (String[])paramsArray[4];
			postIDs = (String[])paramsArray[5];
		}

		/// <summary>
		/// SupportByVersion Office 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff861145.aspx
		/// </summary>
		/// <param name="account">string Account</param>
		/// <param name="postID">string PostID</param>
		/// <param name="parentWindow">Int32 ParentWindow</param>
		/// <param name="xHTML">string xHTML</param>
		/// <param name="title">string Title</param>
		/// <param name="datePosted">string DatePosted</param>
		/// <param name="categories">String[] Categories</param>
		[SupportByVersionAttribute("Office", 12,14,15,16)]
		public void Open(string account, string postID, Int32 parentWindow, out string xHTML, out string title, out string datePosted, out String[] categories)
		{
			ParameterModifier[] modifiers = Invoker.CreateParamModifiers(false,false,false,true,true,true,true);
			xHTML = string.Empty;
			title = string.Empty;
			datePosted = string.Empty;
			categories = null;
			object[] paramsArray = Invoker.ValidateParamsArray(account, postID, parentWindow, xHTML, title, datePosted, (object)categories);
			Invoker.Method(this, "Open", paramsArray, modifiers);
			xHTML = (string)paramsArray[3];
			title = (string)paramsArray[4];
			datePosted = (string)paramsArray[5];
			categories = (String[])paramsArray[6];
		}

		/// <summary>
		/// SupportByVersion Office 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff862458.aspx
		/// </summary>
		/// <param name="account">string Account</param>
		/// <param name="parentWindow">Int32 ParentWindow</param>
		/// <param name="document">object Document</param>
		/// <param name="xHTML">string xHTML</param>
		/// <param name="title">string Title</param>
		/// <param name="dateTime">string DateTime</param>
		/// <param name="categories">String[] Categories</param>
		/// <param name="draft">bool Draft</param>
		/// <param name="postID">string PostID</param>
		/// <param name="publishMessage">string PublishMessage</param>
		[SupportByVersionAttribute("Office", 12,14,15,16)]
		public void PublishPost(string account, Int32 parentWindow, object document, string xHTML, string title, string dateTime, String[] categories, bool draft, out string postID, out string publishMessage)
		{
			ParameterModifier[] modifiers = Invoker.CreateParamModifiers(false,false,false,false,false,false,false,false,true,true);
			postID = string.Empty;
			publishMessage = string.Empty;
			object[] paramsArray = Invoker.ValidateParamsArray(account, parentWindow, document, xHTML, title, dateTime, (object)categories, draft, postID, publishMessage);
			Invoker.Method(this, "PublishPost", paramsArray, modifiers);
			postID = (string)paramsArray[8];
			publishMessage = (string)paramsArray[9];
		}

		/// <summary>
		/// SupportByVersion Office 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff860616.aspx
		/// </summary>
		/// <param name="account">string Account</param>
		/// <param name="parentWindow">Int32 ParentWindow</param>
		/// <param name="document">object Document</param>
		/// <param name="postID">string PostID</param>
		/// <param name="xHTML">string xHTML</param>
		/// <param name="title">string Title</param>
		/// <param name="dateTime">string DateTime</param>
		/// <param name="categories">String[] Categories</param>
		/// <param name="draft">bool Draft</param>
		/// <param name="publishMessage">string PublishMessage</param>
		[SupportByVersionAttribute("Office", 12,14,15,16)]
		public void RepublishPost(string account, Int32 parentWindow, object document, string postID, string xHTML, string title, string dateTime, String[] categories, bool draft, out string publishMessage)
		{
			ParameterModifier[] modifiers = Invoker.CreateParamModifiers(false,false,false,false,false,false,false,false,false,true);
			publishMessage = string.Empty;
			object[] paramsArray = Invoker.ValidateParamsArray(account, parentWindow, document, postID, xHTML, title, dateTime, (object)categories, draft, publishMessage);
			Invoker.Method(this, "RepublishPost", paramsArray, modifiers);
			publishMessage = (string)paramsArray[9];
		}

		/// <summary>
		/// SupportByVersion Office 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff865355.aspx
		/// </summary>
		/// <param name="account">string Account</param>
		/// <param name="parentWindow">Int32 ParentWindow</param>
		/// <param name="document">object Document</param>
		/// <param name="categories">String[] Categories</param>
		[SupportByVersionAttribute("Office", 12,14,15,16)]
		public void GetCategories(string account, Int32 parentWindow, object document, out String[] categories)
		{
			ParameterModifier[] modifiers = Invoker.CreateParamModifiers(false,false,false,true);
			categories = null;
			object[] paramsArray = Invoker.ValidateParamsArray(account, parentWindow, document, (object)categories);
			Invoker.Method(this, "GetCategories", paramsArray, modifiers);
			categories = (String[])paramsArray[3];
		}

		#endregion
		#pragma warning restore
	}
}