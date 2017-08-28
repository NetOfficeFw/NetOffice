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
	[SupportByVersion("Office", 12,14,15,16)]
	[EntityType(EntityType.IsDispatchInterface)]
 	public class IBlogPictureExtensibility : COMObject
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
                    _type = typeof(IBlogPictureExtensibility);
                return _type;
            }
        }
        
        #endregion
        
		#region Ctor

		/// <param name="factory">current used factory core</param>
		/// <param name="parentObject">object there has created the proxy</param>
		/// <param name="proxyShare">proxy share instead if com proxy</param>
		public IBlogPictureExtensibility(Core factory, ICOMObject parentObject, COMProxyShare proxyShare) : base(factory, parentObject, proxyShare)
		{
		}

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
		
		/// <param name="progId">registered progID</param>
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
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff860839.aspx </remarks>
		/// <param name="blogPictureProvider">string blogPictureProvider</param>
		/// <param name="friendlyName">string friendlyName</param>
		[SupportByVersion("Office", 12,14,15,16)]
		public void BlogPictureProviderProperties(out string blogPictureProvider, out string friendlyName)
		{
			ParameterModifier[] modifiers = Invoker.CreateParamModifiers(true,true);
			blogPictureProvider = string.Empty;
			friendlyName = string.Empty;
			object[] paramsArray = Invoker.ValidateParamsArray(blogPictureProvider, friendlyName);
			Invoker.Method(this, "BlogPictureProviderProperties", paramsArray, modifiers);
			blogPictureProvider = paramsArray[0] as string;
			friendlyName = paramsArray[1] as string;
		}

		/// <summary>
		/// SupportByVersion Office 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff862798.aspx </remarks>
		/// <param name="account">string account</param>
		/// <param name="blogProvider">string blogProvider</param>
		/// <param name="parentWindow">Int32 parentWindow</param>
		/// <param name="document">object document</param>
		[SupportByVersion("Office", 12,14,15,16)]
		public void CreatePictureAccount(string account, string blogProvider, Int32 parentWindow, object document)
		{
			 Factory.ExecuteMethod(this, "CreatePictureAccount", account, blogProvider, parentWindow, document);
		}

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
		[SupportByVersion("Office", 12,14,15,16)]
		public void PublishPicture(string account, Int32 parentWindow, object document, object image, out string pictureURI, Int32 imageType)
		{
			ParameterModifier[] modifiers = Invoker.CreateParamModifiers(false,false,false,false,true,false);
			pictureURI = string.Empty;
			object[] paramsArray = Invoker.ValidateParamsArray(account, parentWindow, document, image, pictureURI, imageType);
			Invoker.Method(this, "PublishPicture", paramsArray, modifiers);
			pictureURI = paramsArray[4] as string;
		}

		#endregion

		#pragma warning restore
	}
}
