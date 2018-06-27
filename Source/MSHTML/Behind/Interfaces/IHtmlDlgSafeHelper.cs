using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;
using NetOffice.MSHTMLApi;

namespace NetOffice.MSHTMLApi.Behind
{
	/// <summary>
	/// Interface IHtmlDlgSafeHelper 
	/// SupportByVersion MSHTML, 4
	/// </summary>
	[SupportByVersion("MSHTML", 4)]
	[EntityType(EntityType.IsInterface), BaseType]
 	public class IHtmlDlgSafeHelper : COMObject, NetOffice.MSHTMLApi.IHtmlDlgSafeHelper
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
                    _contractType = typeof(NetOffice.MSHTMLApi.IHtmlDlgSafeHelper);
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
                    _type = typeof(IHtmlDlgSafeHelper);
                return _type;
            }
        }
        
        #endregion
        
		#region Ctor

		/// <summary>
		/// Stub Ctor, not indented to use
		/// </summary>
		public IHtmlDlgSafeHelper() : base()
		{

		}

		#endregion
		
		#region Properties

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		[SupportByVersion("MSHTML", 4), ProxyResult]
		public virtual object fonts
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteReferencePropertyGet(this, "fonts");
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		[SupportByVersion("MSHTML", 4), ProxyResult]
		public virtual object BlockFormats
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteReferencePropertyGet(this, "BlockFormats");
			}
		}

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="initColor">optional object initColor</param>
		[SupportByVersion("MSHTML", 4)]
		public virtual object choosecolordlg(object initColor)
		{
			return InvokerService.InvokeInternal.ExecuteVariantMethodGet(this, "choosecolordlg", initColor);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[CustomMethod]
		[SupportByVersion("MSHTML", 4)]
		public virtual object choosecolordlg()
		{
			return InvokerService.InvokeInternal.ExecuteVariantMethodGet(this, "choosecolordlg");
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="fontName">string fontName</param>
		[SupportByVersion("MSHTML", 4)]
		public virtual object getCharset(string fontName)
		{
			return InvokerService.InvokeInternal.ExecuteVariantMethodGet(this, "getCharset", fontName);
		}

		#endregion

		#pragma warning restore
	}
}

