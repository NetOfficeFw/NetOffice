using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;
using NetOffice.MSHTMLApi;

namespace NetOffice.MSHTMLApi.Behind
{
	/// <summary>
	/// DispatchInterface DispHTMLStyleSheet 
	/// SupportByVersion MSHTML, 4
	/// </summary>
	[SupportByVersion("MSHTML", 4)]
	[EntityType(EntityType.IsDispatchInterface), BaseType]
 	public class DispHTMLStyleSheet : COMObject, NetOffice.MSHTMLApi.DispHTMLStyleSheet
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
                    _contractType = typeof(NetOffice.MSHTMLApi.DispHTMLStyleSheet);
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
                    _type = typeof(DispHTMLStyleSheet);
                return _type;
            }
        }
        
        #endregion
        
		#region Ctor

		/// <summary>
		/// Stub Ctor, not indented to use
		/// </summary>
		public DispHTMLStyleSheet() : base()
		{

		}

		#endregion
		
		#region Properties

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public virtual string title
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "title");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "title", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		[BaseResult]
		public virtual NetOffice.MSHTMLApi.IHTMLStyleSheet parentStyleSheet
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteBaseReferencePropertyGet<NetOffice.MSHTMLApi.IHTMLStyleSheet>(this, "parentStyleSheet");
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		[BaseResult]
		public virtual NetOffice.MSHTMLApi.IHTMLElement owningElement
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteBaseReferencePropertyGet<NetOffice.MSHTMLApi.IHTMLElement>(this, "owningElement");
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public virtual bool disabled
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "disabled");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "disabled", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public virtual bool readOnly
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "readOnly");
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public virtual NetOffice.MSHTMLApi.IHTMLStyleSheetsCollection imports
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.MSHTMLApi.IHTMLStyleSheetsCollection>(this, "imports", typeof(NetOffice.MSHTMLApi.IHTMLStyleSheetsCollection));
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public virtual string href
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "href");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "href", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public virtual string type
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "type");
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public virtual string id
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "id");
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public virtual string media
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "media");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "media", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public virtual string cssText
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "cssText");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "cssText", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public virtual NetOffice.MSHTMLApi.IHTMLStyleSheetRulesCollection rules
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.MSHTMLApi.IHTMLStyleSheetRulesCollection>(this, "rules", typeof(NetOffice.MSHTMLApi.IHTMLStyleSheetRulesCollection));
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public virtual NetOffice.MSHTMLApi.IHTMLStyleSheetPagesCollection pages
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.MSHTMLApi.IHTMLStyleSheetPagesCollection>(this, "pages", typeof(NetOffice.MSHTMLApi.IHTMLStyleSheetPagesCollection));
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public virtual string ie8_href
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "ie8_href");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "ie8_href", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public virtual bool isAlternate
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "isAlternate");
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public virtual bool isPrefAlternate
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "isPrefAlternate");
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		[SupportByVersion("MSHTML", 4), ProxyResult]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public virtual object constructor
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteReferencePropertyGet(this, "constructor");
			}
		}

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="bstrUrl">string bstrUrl</param>
		/// <param name="lIndex">optional Int32 lIndex = -1</param>
		[SupportByVersion("MSHTML", 4)]
		public virtual Int32 addImport(string bstrUrl, object lIndex)
		{
			return InvokerService.InvokeInternal.ExecuteInt32MethodGet(this, "addImport", bstrUrl, lIndex);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="bstrUrl">string bstrUrl</param>
		[CustomMethod]
		[SupportByVersion("MSHTML", 4)]
		public virtual Int32 addImport(string bstrUrl)
		{
			return InvokerService.InvokeInternal.ExecuteInt32MethodGet(this, "addImport", bstrUrl);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="bstrSelector">string bstrSelector</param>
		/// <param name="bstrStyle">string bstrStyle</param>
		/// <param name="lIndex">optional Int32 lIndex = -1</param>
		[SupportByVersion("MSHTML", 4)]
		public virtual Int32 addRule(string bstrSelector, string bstrStyle, object lIndex)
		{
			return InvokerService.InvokeInternal.ExecuteInt32MethodGet(this, "addRule", bstrSelector, bstrStyle, lIndex);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="bstrSelector">string bstrSelector</param>
		/// <param name="bstrStyle">string bstrStyle</param>
		[CustomMethod]
		[SupportByVersion("MSHTML", 4)]
		public virtual Int32 addRule(string bstrSelector, string bstrStyle)
		{
			return InvokerService.InvokeInternal.ExecuteInt32MethodGet(this, "addRule", bstrSelector, bstrStyle);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="lIndex">Int32 lIndex</param>
		[SupportByVersion("MSHTML", 4)]
		public virtual void removeImport(Int32 lIndex)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "removeImport", lIndex);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="lIndex">Int32 lIndex</param>
		[SupportByVersion("MSHTML", 4)]
		public virtual void removeRule(Int32 lIndex)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "removeRule", lIndex);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="bstrSelector">string bstrSelector</param>
		/// <param name="bstrStyle">string bstrStyle</param>
		/// <param name="lIndex">optional Int32 lIndex = -1</param>
		[SupportByVersion("MSHTML", 4)]
		public virtual Int32 addPageRule(string bstrSelector, string bstrStyle, object lIndex)
		{
			return InvokerService.InvokeInternal.ExecuteInt32MethodGet(this, "addPageRule", bstrSelector, bstrStyle, lIndex);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="bstrSelector">string bstrSelector</param>
		/// <param name="bstrStyle">string bstrStyle</param>
		[CustomMethod]
		[SupportByVersion("MSHTML", 4)]
		public virtual Int32 addPageRule(string bstrSelector, string bstrStyle)
		{
			return InvokerService.InvokeInternal.ExecuteInt32MethodGet(this, "addPageRule", bstrSelector, bstrStyle);
		}

		#endregion

		#pragma warning restore
	}
}


