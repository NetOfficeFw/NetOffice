using System.Reflection;
using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;
using NetOffice.OfficeApi;

namespace NetOffice.OfficeApi.Behind
{
    /// <summary>
    /// DispatchInterface IBlogPictureExtensibility 
    /// SupportByVersion Office, 12,14,15,16
    /// </summary>
    /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff860265.aspx </remarks>
    [SupportByVersion("Office", 12, 14, 15, 16)]
    [EntityType(EntityType.IsDispatchInterface)]
    public class IBlogPictureExtensibility : COMObject, NetOffice.OfficeApi.IBlogPictureExtensibility
    {
        #pragma warning disable

        #region Type Information

        /// <summary>
        /// Contract Type
        /// </summary>
        [EditorBrowsable(EditorBrowsableState.Advanced), Browsable(false), Category("NetOffice"), CoreOverridden]
        public override Type ContractType
        {
            get
            {
                if(null == _contractType)
                    _contractType = typeof(NetOffice.OfficeApi.IBlogPictureExtensibility);
                return _contractType;
            }
        }
        private static Type _contractType;


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

		/// <summary>
		/// Stub Ctor, not indented to use
		/// </summary>
		public IBlogPictureExtensibility() : base()
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
        [SupportByVersion("Office", 12, 14, 15, 16)]
        public virtual void BlogPictureProviderProperties(out string blogPictureProvider, out string friendlyName)
        {
            ParameterModifier[] modifiers = Invoker.CreateParamModifiers(true, true);
            blogPictureProvider = string.Empty;
            friendlyName = string.Empty;
            object[] paramsArray = new object[] { blogPictureProvider, friendlyName };

            InvokerService.InvokeInternal.ExecuteMethodExtended(this, "BlogPictureProviderProperties", paramsArray, modifiers);
            
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
        [SupportByVersion("Office", 12, 14, 15, 16)]
        public virtual void CreatePictureAccount(string account, string blogProvider, Int32 parentWindow, object document)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "CreatePictureAccount", account, blogProvider, parentWindow, document);
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
        [SupportByVersion("Office", 12, 14, 15, 16)]
        public virtual void PublishPicture(string account, Int32 parentWindow, object document, object image, out string pictureURI, Int32 imageType)
        {
            ParameterModifier[] modifiers = Invoker.CreateParamModifiers(false, false, false, false, true, false);
            pictureURI = string.Empty;
            object[] paramsArray = new object[] { account, parentWindow, document, image, pictureURI, imageType };
            InvokerService.InvokeInternal.ExecuteMethod(this, "PublishPicture", paramsArray, modifiers);
            pictureURI = paramsArray[4] as string;
        }

        #endregion

        #pragma warning restore
    }
}
