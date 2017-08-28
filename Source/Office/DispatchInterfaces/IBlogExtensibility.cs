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
	[SupportByVersion("Office", 12,14,15,16)]
	[EntityType(EntityType.IsDispatchInterface)]
 	public class IBlogExtensibility : COMObject
	{
		#pragma warning disable

		#region Type Information

		/// <summary>
		/// Instance Type
		/// </summary>
		[EditorBrowsable(EditorBrowsableState.Advanced), Browsable(false), Category("NetOffice"), CoreOverridden]
		public override Type InstanceType
		{
			get
			{
				return LateBindingApiWrapperType;
			}
		}

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
        
		#region Ctor

		/// <param name="factory">current used factory core</param>
		/// <param name="parentObject">object there has created the proxy</param>
		/// <param name="proxyShare">proxy share instead if com proxy</param>
		public IBlogExtensibility(Core factory, ICOMObject parentObject, COMProxyShare proxyShare) : base(factory, parentObject, proxyShare)
		{
		}

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
		
		/// <param name="progId">registered progID</param>
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
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff862840.aspx </remarks>
		/// <param name="blogProvider">string blogProvider</param>
		/// <param name="friendlyName">string friendlyName</param>
		/// <param name="categorySupport">NetOffice.OfficeApi.Enums.MsoBlogCategorySupport categorySupport</param>
		/// <param name="padding">bool padding</param>
		[SupportByVersion("Office", 12,14,15,16)]
		public void BlogProviderProperties(out string blogProvider, out string friendlyName, out NetOffice.OfficeApi.Enums.MsoBlogCategorySupport categorySupport, out bool padding)
		{
			ParameterModifier[] modifiers = Invoker.CreateParamModifiers(true,true,true,true);
			blogProvider = string.Empty;
			friendlyName = string.Empty;
			categorySupport = 0;
			padding = false;
			object[] paramsArray = Invoker.ValidateParamsArray(blogProvider, friendlyName, categorySupport, padding);
			Invoker.Method(this, "BlogProviderProperties", paramsArray, modifiers);
			blogProvider = paramsArray[0] as string;
			friendlyName = paramsArray[1] as string;
			categorySupport = (NetOffice.OfficeApi.Enums.MsoBlogCategorySupport)paramsArray[2];
			padding = (bool)paramsArray[3];
		}

		/// <summary>
		/// SupportByVersion Office 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff863154.aspx </remarks>
		/// <param name="account">string account</param>
		/// <param name="parentWindow">Int32 parentWindow</param>
		/// <param name="document">object document</param>
		/// <param name="newAccount">bool newAccount</param>
		/// <param name="showPictureUI">bool showPictureUI</param>
		[SupportByVersion("Office", 12,14,15,16)]
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
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff860220.aspx </remarks>
		/// <param name="account">string account</param>
		/// <param name="parentWindow">Int32 parentWindow</param>
		/// <param name="document">object document</param>
		/// <param name="blogNames">String[] blogNames</param>
		/// <param name="blogIDs">String[] blogIDs</param>
		/// <param name="blogURLs">String[] blogURLs</param>
		[SupportByVersion("Office", 12,14,15,16)]
		public void GetUserBlogs(string account, Int32 parentWindow, object document, out String[] blogNames, out String[] blogIDs, out String[] blogURLs)
		{
			ParameterModifier[] modifiers = Invoker.CreateParamModifiers(false,false,false,true,true,true);
			blogNames = null;
			blogIDs = null;
			blogURLs = null;
			object[] paramsArray = Invoker.ValidateParamsArray(account, parentWindow, document, (object)blogNames, (object)blogIDs, (object)blogURLs);
			Invoker.Method(this, "GetUserBlogs", paramsArray, modifiers);
			blogNames = paramsArray[3] as String[];
			blogIDs = paramsArray[4] as String[];
            blogURLs = paramsArray[5] as String[];
        }

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
		[SupportByVersion("Office", 12,14,15,16)]
		public void GetRecentPosts(string account, Int32 parentWindow, object document, out String[] postTitles, out String[] postDates, out String[] postIDs)
		{
			ParameterModifier[] modifiers = Invoker.CreateParamModifiers(false,false,false,true,true,true);
			postTitles = null;
			postDates = null;
			postIDs = null;
			object[] paramsArray = Invoker.ValidateParamsArray(account, parentWindow, document, (object)postTitles, (object)postDates, (object)postIDs);
			Invoker.Method(this, "GetRecentPosts", paramsArray, modifiers);
			postTitles = paramsArray[3] as String[];
			postDates = paramsArray[4] as String[];
			postIDs = paramsArray[5] as String[];
		}

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
		[SupportByVersion("Office", 12,14,15,16)]
		public void Open(string account, string postID, Int32 parentWindow, out string xHTML, out string title, out string datePosted, out String[] categories)
		{
			ParameterModifier[] modifiers = Invoker.CreateParamModifiers(false,false,false,true,true,true,true);
			xHTML = string.Empty;
			title = string.Empty;
			datePosted = string.Empty;
			categories = null;
			object[] paramsArray = Invoker.ValidateParamsArray(account, postID, parentWindow, xHTML, title, datePosted, (object)categories);
			Invoker.Method(this, "Open", paramsArray, modifiers);
			xHTML = paramsArray[3] as string;
			title = paramsArray[4] as string;
			datePosted = paramsArray[5] as string;
			categories = paramsArray[6] as String[];
		}

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
		[SupportByVersion("Office", 12,14,15,16)]
		public void PublishPost(string account, Int32 parentWindow, object document, string xHTML, string title, string dateTime, String[] categories, bool draft, out string postID, out string publishMessage)
		{
			ParameterModifier[] modifiers = Invoker.CreateParamModifiers(false,false,false,false,false,false,false,false,true,true);
			postID = string.Empty;
			publishMessage = string.Empty;
			object[] paramsArray = Invoker.ValidateParamsArray(account, parentWindow, document, xHTML, title, dateTime, (object)categories, draft, postID, publishMessage);
			Invoker.Method(this, "PublishPost", paramsArray, modifiers);
			postID = paramsArray[8] as string;
			publishMessage = paramsArray[9] as string;
		}

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
		[SupportByVersion("Office", 12,14,15,16)]
		public void RepublishPost(string account, Int32 parentWindow, object document, string postID, string xHTML, string title, string dateTime, String[] categories, bool draft, out string publishMessage)
		{
			ParameterModifier[] modifiers = Invoker.CreateParamModifiers(false,false,false,false,false,false,false,false,false,true);
			publishMessage = string.Empty;
			object[] paramsArray = Invoker.ValidateParamsArray(account, parentWindow, document, postID, xHTML, title, dateTime, (object)categories, draft, publishMessage);
			Invoker.Method(this, "RepublishPost", paramsArray, modifiers);
			publishMessage = paramsArray[9] as string;
		}

		/// <summary>
		/// SupportByVersion Office 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff865355.aspx </remarks>
		/// <param name="account">string account</param>
		/// <param name="parentWindow">Int32 parentWindow</param>
		/// <param name="document">object document</param>
		/// <param name="categories">String[] categories</param>
		[SupportByVersion("Office", 12,14,15,16)]
		public void GetCategories(string account, Int32 parentWindow, object document, out String[] categories)
		{
			ParameterModifier[] modifiers = Invoker.CreateParamModifiers(false,false,false,true);
			categories = null;
			object[] paramsArray = Invoker.ValidateParamsArray(account, parentWindow, document, (object)categories);
			Invoker.Method(this, "GetCategories", paramsArray, modifiers);
			categories = paramsArray[3] as String[];
		}

		#endregion

		#pragma warning restore
	}
}
