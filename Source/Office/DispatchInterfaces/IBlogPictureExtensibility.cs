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
	/// DispatchInterface IBlogPictureExtensibility 
	/// SupportByVersion Office, 12,14,15,16
	/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff860265.aspx
	///</summary>
	[SupportByVersionAttribute("Office", 12,14,15,16)]
	[EntityTypeAttribute(EntityType.IsDispatchInterface)]
	public class IBlogPictureExtensibility : COMObject
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
                    _type = typeof(IBlogPictureExtensibility);
                    
                return _type;
            }
        }
        
        #endregion
        
		#region Construction

		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
		public IBlogPictureExtensibility(Core factory, ICOMObject parentObject, object comProxy) : base(factory, parentObject, comProxy)
		{
			
		}

        ///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public IBlogPictureExtensibility(ICOMObject parentObject, object comProxy) : base(parentObject, comProxy)
		{
		}
		
		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public IBlogPictureExtensibility(Core factory, ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(factory, parentObject, comProxy, comProxyType)
		{

		}

		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public IBlogPictureExtensibility(ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(parentObject, comProxy, comProxyType)
		{
		}
		
		///<param name="replacedObject">object to replaced. replacedObject are not usable after this action</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public IBlogPictureExtensibility(ICOMObject replacedObject) : base(replacedObject)
		{
		}
		
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public IBlogPictureExtensibility() : base()
		{
		}
		
		/// <param name="progId">registered ProgID</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public IBlogPictureExtensibility(string progId) : base(progId)
		{
		}
		
		#endregion
		
		#region Properties

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion Office 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff860839.aspx
		/// </summary>
		/// <param name="blogPictureProvider">string BlogPictureProvider</param>
		/// <param name="friendlyName">string FriendlyName</param>
		[SupportByVersionAttribute("Office", 12,14,15,16)]
		public void BlogPictureProviderProperties(out string blogPictureProvider, out string friendlyName)
		{
			ParameterModifier[] modifiers = Invoker.CreateParamModifiers(true,true);
			blogPictureProvider = string.Empty;
			friendlyName = string.Empty;
			object[] paramsArray = Invoker.ValidateParamsArray(blogPictureProvider, friendlyName);
			Invoker.Method(this, "BlogPictureProviderProperties", paramsArray, modifiers);
			blogPictureProvider = (string)paramsArray[0];
			friendlyName = (string)paramsArray[1];
		}

		/// <summary>
		/// SupportByVersion Office 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff862798.aspx
		/// </summary>
		/// <param name="account">string Account</param>
		/// <param name="blogProvider">string BlogProvider</param>
		/// <param name="parentWindow">Int32 ParentWindow</param>
		/// <param name="document">object Document</param>
		[SupportByVersionAttribute("Office", 12,14,15,16)]
		public void CreatePictureAccount(string account, string blogProvider, Int32 parentWindow, object document)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(account, blogProvider, parentWindow, document);
			Invoker.Method(this, "CreatePictureAccount", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Office 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff864012.aspx
		/// </summary>
		/// <param name="account">string Account</param>
		/// <param name="parentWindow">Int32 ParentWindow</param>
		/// <param name="document">object Document</param>
		/// <param name="image">object Image</param>
		/// <param name="pictureURI">string PictureURI</param>
		/// <param name="imageType">Int32 ImageType</param>
		[SupportByVersionAttribute("Office", 12,14,15,16)]
		public void PublishPicture(string account, Int32 parentWindow, object document, object image, out string pictureURI, Int32 imageType)
		{
			ParameterModifier[] modifiers = Invoker.CreateParamModifiers(false,false,false,false,true,false);
			pictureURI = string.Empty;
			object[] paramsArray = Invoker.ValidateParamsArray(account, parentWindow, document, image, pictureURI, imageType);
			Invoker.Method(this, "PublishPicture", paramsArray, modifiers);
			pictureURI = (string)paramsArray[4];
		}

		#endregion
		#pragma warning restore
	}
}