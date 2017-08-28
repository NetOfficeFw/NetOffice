using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.MSHTMLApi
{
	/// <summary>
	/// DispatchInterface IHTMLStyleSheet 
	/// SupportByVersion MSHTML, 4
	/// </summary>
	[SupportByVersion("MSHTML", 4)]
	[EntityType(EntityType.IsDispatchInterface)]
 	public class IHTMLStyleSheet : DispHTMLStyleSheet
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
                    _type = typeof(IHTMLStyleSheet);
                return _type;
            }
        }
        
        #endregion
        
		#region Ctor

		/// <param name="factory">current used factory core</param>
		/// <param name="parentObject">object there has created the proxy</param>
		/// <param name="proxyShare">proxy share instead if com proxy</param>
		public IHTMLStyleSheet(Core factory, ICOMObject parentObject, COMProxyShare proxyShare) : base(factory, parentObject, proxyShare)
		{
		}

		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
		public IHTMLStyleSheet(Core factory, ICOMObject parentObject, object comProxy) : base(factory, parentObject, comProxy)
		{
			
		}

        ///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public IHTMLStyleSheet(ICOMObject parentObject, object comProxy) : base(parentObject, comProxy)
		{
		}
		
		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public IHTMLStyleSheet(Core factory, ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(factory, parentObject, comProxy, comProxyType)
		{

		}

		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public IHTMLStyleSheet(ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(parentObject, comProxy, comProxyType)
		{
		}
		
		///<param name="replacedObject">object to replaced. replacedObject are not usable after this action</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public IHTMLStyleSheet(ICOMObject replacedObject) : base(replacedObject)
		{
		}
		
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public IHTMLStyleSheet() : base()
		{
		}
		
		/// <param name="progId">registered progID</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public IHTMLStyleSheet(string progId) : base(progId)
		{
		}
		
		#endregion
		
		#region Properties

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public string title
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "title");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "title", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		[BaseResult]
		public NetOffice.MSHTMLApi.IHTMLStyleSheet parentStyleSheet
		{
			get
			{
				return Factory.ExecuteBaseReferencePropertyGet<NetOffice.MSHTMLApi.IHTMLStyleSheet>(this, "parentStyleSheet");
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		[BaseResult]
		public NetOffice.MSHTMLApi.IHTMLElement owningElement
		{
			get
			{
				return Factory.ExecuteBaseReferencePropertyGet<NetOffice.MSHTMLApi.IHTMLElement>(this, "owningElement");
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public bool disabled
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "disabled");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "disabled", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public bool readOnly
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "readOnly");
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public NetOffice.MSHTMLApi.IHTMLStyleSheetsCollection imports
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.MSHTMLApi.IHTMLStyleSheetsCollection>(this, "imports", NetOffice.MSHTMLApi.IHTMLStyleSheetsCollection.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public string href
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "href");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "href", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public string type
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "type");
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public string id
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "id");
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public string media
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "media");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "media", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public string cssText
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "cssText");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "cssText", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public NetOffice.MSHTMLApi.IHTMLStyleSheetRulesCollection rules
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.MSHTMLApi.IHTMLStyleSheetRulesCollection>(this, "rules", NetOffice.MSHTMLApi.IHTMLStyleSheetRulesCollection.LateBindingApiWrapperType);
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
		public Int32 addImport(string bstrUrl, object lIndex)
		{
			return Factory.ExecuteInt32MethodGet(this, "addImport", bstrUrl, lIndex);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="bstrUrl">string bstrUrl</param>
		[CustomMethod]
		[SupportByVersion("MSHTML", 4)]
		public Int32 addImport(string bstrUrl)
		{
			return Factory.ExecuteInt32MethodGet(this, "addImport", bstrUrl);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="bstrSelector">string bstrSelector</param>
		/// <param name="bstrStyle">string bstrStyle</param>
		/// <param name="lIndex">optional Int32 lIndex = -1</param>
		[SupportByVersion("MSHTML", 4)]
		public Int32 addRule(string bstrSelector, string bstrStyle, object lIndex)
		{
			return Factory.ExecuteInt32MethodGet(this, "addRule", bstrSelector, bstrStyle, lIndex);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="bstrSelector">string bstrSelector</param>
		/// <param name="bstrStyle">string bstrStyle</param>
		[CustomMethod]
		[SupportByVersion("MSHTML", 4)]
		public Int32 addRule(string bstrSelector, string bstrStyle)
		{
			return Factory.ExecuteInt32MethodGet(this, "addRule", bstrSelector, bstrStyle);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="lIndex">Int32 lIndex</param>
		[SupportByVersion("MSHTML", 4)]
		public void removeImport(Int32 lIndex)
		{
			 Factory.ExecuteMethod(this, "removeImport", lIndex);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="lIndex">Int32 lIndex</param>
		[SupportByVersion("MSHTML", 4)]
		public void removeRule(Int32 lIndex)
		{
			 Factory.ExecuteMethod(this, "removeRule", lIndex);
		}

		#endregion

		#pragma warning restore
	}
}
