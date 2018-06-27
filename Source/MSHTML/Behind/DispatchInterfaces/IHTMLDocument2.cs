using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;
using NetOffice.MSHTMLApi;

namespace NetOffice.MSHTMLApi.Behind
{
	/// <summary>
	/// DispatchInterface IHTMLDocument2 
	/// SupportByVersion MSHTML, 4
	/// </summary>
	[SupportByVersion("MSHTML", 4)]
	[EntityType(EntityType.IsDispatchInterface)]
 	public class IHTMLDocument2 : IHTMLDocument, NetOffice.MSHTMLApi.IHTMLDocument2
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
                    _contractType = typeof(NetOffice.MSHTMLApi.IHTMLDocument2);
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
                    _type = typeof(IHTMLDocument2);
                return _type;
            }
        }
        
        #endregion
        
		#region Ctor

		/// <summary>
		/// Stub Ctor, not indented to use
		/// </summary>
		public IHTMLDocument2() : base()
		{

		}

		#endregion
		
		#region Properties

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		[BaseResult]
		public virtual NetOffice.MSHTMLApi.IHTMLElementCollection all
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteBaseReferencePropertyGet<NetOffice.MSHTMLApi.IHTMLElementCollection>(this, "all");
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		[BaseResult]
		public virtual NetOffice.MSHTMLApi.IHTMLElement body
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteBaseReferencePropertyGet<NetOffice.MSHTMLApi.IHTMLElement>(this, "body");
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		[BaseResult]
		public virtual NetOffice.MSHTMLApi.IHTMLElement activeElement
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteBaseReferencePropertyGet<NetOffice.MSHTMLApi.IHTMLElement>(this, "activeElement");
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		[BaseResult]
		public virtual NetOffice.MSHTMLApi.IHTMLElementCollection images
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteBaseReferencePropertyGet<NetOffice.MSHTMLApi.IHTMLElementCollection>(this, "images");
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		[BaseResult]
		public virtual NetOffice.MSHTMLApi.IHTMLElementCollection applets
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteBaseReferencePropertyGet<NetOffice.MSHTMLApi.IHTMLElementCollection>(this, "applets");
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		[BaseResult]
		public virtual NetOffice.MSHTMLApi.IHTMLElementCollection links
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteBaseReferencePropertyGet<NetOffice.MSHTMLApi.IHTMLElementCollection>(this, "links");
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		[BaseResult]
		public virtual NetOffice.MSHTMLApi.IHTMLElementCollection forms
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteBaseReferencePropertyGet<NetOffice.MSHTMLApi.IHTMLElementCollection>(this, "forms");
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		[BaseResult]
		public virtual NetOffice.MSHTMLApi.IHTMLElementCollection anchors
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteBaseReferencePropertyGet<NetOffice.MSHTMLApi.IHTMLElementCollection>(this, "anchors");
			}
		}

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
		public virtual NetOffice.MSHTMLApi.IHTMLElementCollection scripts
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteBaseReferencePropertyGet<NetOffice.MSHTMLApi.IHTMLElementCollection>(this, "scripts");
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public virtual string designMode
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "designMode");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "designMode", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public virtual NetOffice.MSHTMLApi.IHTMLSelectionObject selection
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.MSHTMLApi.IHTMLSelectionObject>(this, "selection", typeof(NetOffice.MSHTMLApi.IHTMLSelectionObject));
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public virtual string readyState
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "readyState");
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		[BaseResult]
		public virtual NetOffice.MSHTMLApi.IHTMLFramesCollection2 frames
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteBaseReferencePropertyGet<NetOffice.MSHTMLApi.IHTMLFramesCollection2>(this, "frames");
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		[BaseResult]
		public virtual NetOffice.MSHTMLApi.IHTMLElementCollection embeds
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteBaseReferencePropertyGet<NetOffice.MSHTMLApi.IHTMLElementCollection>(this, "embeds");
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		[BaseResult]
		public virtual NetOffice.MSHTMLApi.IHTMLElementCollection plugins
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteBaseReferencePropertyGet<NetOffice.MSHTMLApi.IHTMLElementCollection>(this, "plugins");
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public virtual object alinkColor
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "alinkColor");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteVariantPropertySet(this, "alinkColor", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public virtual object bgColor
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "bgColor");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteVariantPropertySet(this, "bgColor", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public virtual object fgColor
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "fgColor");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteVariantPropertySet(this, "fgColor", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public virtual object linkColor
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "linkColor");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteVariantPropertySet(this, "linkColor", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public virtual object vlinkColor
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "vlinkColor");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteVariantPropertySet(this, "vlinkColor", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public virtual string referrer
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "referrer");
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public virtual NetOffice.MSHTMLApi.IHTMLLocation location
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.MSHTMLApi.IHTMLLocation>(this, "location", typeof(NetOffice.MSHTMLApi.IHTMLLocation));
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public virtual string lastModified
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "lastModified");
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public virtual string url
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "url");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "url", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public virtual string domain
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "domain");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "domain", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public virtual string cookie
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "cookie");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "cookie", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public virtual bool expando
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "expando");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "expando", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public virtual string charset
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "charset");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "charset", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public virtual string defaultCharset
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "defaultCharset");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "defaultCharset", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public virtual string mimeType
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "mimeType");
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public virtual string fileSize
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "fileSize");
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public virtual string fileCreatedDate
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "fileCreatedDate");
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public virtual string fileModifiedDate
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "fileModifiedDate");
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public virtual string fileUpdatedDate
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "fileUpdatedDate");
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public virtual string security
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "security");
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public virtual string protocol
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "protocol");
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public virtual string nameProp
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "nameProp");
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public virtual object onhelp
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "onhelp");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteVariantPropertySet(this, "onhelp", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public virtual object onclick
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "onclick");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteVariantPropertySet(this, "onclick", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public virtual object ondblclick
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "ondblclick");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteVariantPropertySet(this, "ondblclick", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public virtual object onkeyup
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "onkeyup");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteVariantPropertySet(this, "onkeyup", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public virtual object onkeydown
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "onkeydown");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteVariantPropertySet(this, "onkeydown", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public virtual object onkeypress
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "onkeypress");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteVariantPropertySet(this, "onkeypress", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public virtual object onmouseup
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "onmouseup");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteVariantPropertySet(this, "onmouseup", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public virtual object onmousedown
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "onmousedown");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteVariantPropertySet(this, "onmousedown", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public virtual object onmousemove
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "onmousemove");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteVariantPropertySet(this, "onmousemove", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public virtual object onmouseout
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "onmouseout");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteVariantPropertySet(this, "onmouseout", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public virtual object onmouseover
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "onmouseover");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteVariantPropertySet(this, "onmouseover", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public virtual object onreadystatechange
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "onreadystatechange");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteVariantPropertySet(this, "onreadystatechange", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public virtual object onafterupdate
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "onafterupdate");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteVariantPropertySet(this, "onafterupdate", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public virtual object onrowexit
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "onrowexit");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteVariantPropertySet(this, "onrowexit", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public virtual object onrowenter
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "onrowenter");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteVariantPropertySet(this, "onrowenter", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public virtual object ondragstart
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "ondragstart");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteVariantPropertySet(this, "ondragstart", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public virtual object onselectstart
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "onselectstart");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteVariantPropertySet(this, "onselectstart", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public virtual NetOffice.MSHTMLApi.IHTMLWindow2 parentWindow
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.MSHTMLApi.IHTMLWindow2>(this, "parentWindow", typeof(NetOffice.MSHTMLApi.IHTMLWindow2));
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public virtual NetOffice.MSHTMLApi.IHTMLStyleSheetsCollection styleSheets
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.MSHTMLApi.IHTMLStyleSheetsCollection>(this, "styleSheets", typeof(NetOffice.MSHTMLApi.IHTMLStyleSheetsCollection));
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public virtual object onbeforeupdate
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "onbeforeupdate");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteVariantPropertySet(this, "onbeforeupdate", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public virtual object onerrorupdate
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "onerrorupdate");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteVariantPropertySet(this, "onerrorupdate", value);
			}
		}

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="psarray">optional object[] psarray</param>
		[SupportByVersion("MSHTML", 4)]
		public virtual void write(object[] psarray)
		{
            object[] paramsArray = Invoker.ValidateParamsArray((object)psarray);
            Invoker.Method(this, "write", paramsArray);
        }

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[CustomMethod]
		[SupportByVersion("MSHTML", 4)]
		public virtual void write()
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "write");
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="psarray">optional object[] psarray</param>
		[SupportByVersion("MSHTML", 4)]
		public virtual void writeln(object[] psarray)
		{
            object[] paramsArray = Invoker.ValidateParamsArray((object)psarray);
            Invoker.Method(this, "writeln", paramsArray);
        }

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[CustomMethod]
		[SupportByVersion("MSHTML", 4)]
		public virtual void writeln()
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "writeln");
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="url">optional string url = text/html</param>
		/// <param name="name">optional object name</param>
		/// <param name="features">optional object features</param>
		/// <param name="replace">optional object replace</param>
		[SupportByVersion("MSHTML", 4)]
		public virtual object open(object url, object name, object features, object replace)
		{
			return InvokerService.InvokeInternal.ExecuteVariantMethodGet(this, "open", url, name, features, replace);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[CustomMethod]
		[SupportByVersion("MSHTML", 4)]
		public virtual object open()
		{
			return InvokerService.InvokeInternal.ExecuteVariantMethodGet(this, "open");
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="url">optional string url = text/html</param>
		[CustomMethod]
		[SupportByVersion("MSHTML", 4)]
		public virtual object open(object url)
		{
			return InvokerService.InvokeInternal.ExecuteVariantMethodGet(this, "open", url);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="url">optional string url = text/html</param>
		/// <param name="name">optional object name</param>
		[CustomMethod]
		[SupportByVersion("MSHTML", 4)]
		public virtual object open(object url, object name)
		{
			return InvokerService.InvokeInternal.ExecuteVariantMethodGet(this, "open", url, name);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="url">optional string url = text/html</param>
		/// <param name="name">optional object name</param>
		/// <param name="features">optional object features</param>
		[CustomMethod]
		[SupportByVersion("MSHTML", 4)]
		public virtual object open(object url, object name, object features)
		{
			return InvokerService.InvokeInternal.ExecuteVariantMethodGet(this, "open", url, name, features);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public virtual void close()
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "close");
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public virtual void clear()
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "clear");
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="cmdID">string cmdID</param>
		[SupportByVersion("MSHTML", 4)]
		public virtual bool queryCommandSupported(string cmdID)
		{
			return InvokerService.InvokeInternal.ExecuteBoolMethodGet(this, "queryCommandSupported", cmdID);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="cmdID">string cmdID</param>
		[SupportByVersion("MSHTML", 4)]
		public virtual bool queryCommandEnabled(string cmdID)
		{
			return InvokerService.InvokeInternal.ExecuteBoolMethodGet(this, "queryCommandEnabled", cmdID);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="cmdID">string cmdID</param>
		[SupportByVersion("MSHTML", 4)]
		public virtual bool queryCommandState(string cmdID)
		{
			return InvokerService.InvokeInternal.ExecuteBoolMethodGet(this, "queryCommandState", cmdID);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="cmdID">string cmdID</param>
		[SupportByVersion("MSHTML", 4)]
		public virtual bool queryCommandIndeterm(string cmdID)
		{
			return InvokerService.InvokeInternal.ExecuteBoolMethodGet(this, "queryCommandIndeterm", cmdID);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="cmdID">string cmdID</param>
		[SupportByVersion("MSHTML", 4)]
		public virtual string queryCommandText(string cmdID)
		{
			return InvokerService.InvokeInternal.ExecuteStringMethodGet(this, "queryCommandText", cmdID);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="cmdID">string cmdID</param>
		[SupportByVersion("MSHTML", 4)]
		public virtual object queryCommandValue(string cmdID)
		{
			return InvokerService.InvokeInternal.ExecuteVariantMethodGet(this, "queryCommandValue", cmdID);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="cmdID">string cmdID</param>
		/// <param name="showUI">optional bool showUI = false</param>
		/// <param name="value">optional object value</param>
		[SupportByVersion("MSHTML", 4)]
		public virtual bool execCommand(string cmdID, object showUI, object value)
		{
			return InvokerService.InvokeInternal.ExecuteBoolMethodGet(this, "execCommand", cmdID, showUI, value);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="cmdID">string cmdID</param>
		[CustomMethod]
		[SupportByVersion("MSHTML", 4)]
		public virtual bool execCommand(string cmdID)
		{
			return InvokerService.InvokeInternal.ExecuteBoolMethodGet(this, "execCommand", cmdID);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="cmdID">string cmdID</param>
		/// <param name="showUI">optional bool showUI = false</param>
		[CustomMethod]
		[SupportByVersion("MSHTML", 4)]
		public virtual bool execCommand(string cmdID, object showUI)
		{
			return InvokerService.InvokeInternal.ExecuteBoolMethodGet(this, "execCommand", cmdID, showUI);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="cmdID">string cmdID</param>
		[SupportByVersion("MSHTML", 4)]
		public virtual bool execCommandShowHelp(string cmdID)
		{
			return InvokerService.InvokeInternal.ExecuteBoolMethodGet(this, "execCommandShowHelp", cmdID);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="eTag">string eTag</param>
		[SupportByVersion("MSHTML", 4)]
		[BaseResult]
		public virtual NetOffice.MSHTMLApi.IHTMLElement createElement(string eTag)
		{
			return InvokerService.InvokeInternal.ExecuteBaseReferenceMethodGet<NetOffice.MSHTMLApi.IHTMLElement>(this, "createElement", eTag);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="x">Int32 x</param>
		/// <param name="y">Int32 y</param>
		[SupportByVersion("MSHTML", 4)]
		[BaseResult]
		public virtual NetOffice.MSHTMLApi.IHTMLElement elementFromPoint(Int32 x, Int32 y)
		{
			return InvokerService.InvokeInternal.ExecuteBaseReferenceMethodGet<NetOffice.MSHTMLApi.IHTMLElement>(this, "elementFromPoint", x, y);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public virtual string toString()
		{
			return InvokerService.InvokeInternal.ExecuteStringMethodGet(this, "toString");
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="bstrHref">optional string bstrHref = </param>
		/// <param name="lIndex">optional Int32 lIndex = -1</param>
		[SupportByVersion("MSHTML", 4)]
		[BaseResult]
		public virtual NetOffice.MSHTMLApi.IHTMLStyleSheet createStyleSheet(object bstrHref, object lIndex)
		{
			return InvokerService.InvokeInternal.ExecuteBaseReferenceMethodGet<NetOffice.MSHTMLApi.IHTMLStyleSheet>(this, "createStyleSheet", bstrHref, lIndex);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[CustomMethod]
		[BaseResult]
		[SupportByVersion("MSHTML", 4)]
		public virtual NetOffice.MSHTMLApi.IHTMLStyleSheet createStyleSheet()
		{
			return InvokerService.InvokeInternal.ExecuteBaseReferenceMethodGet<NetOffice.MSHTMLApi.IHTMLStyleSheet>(this, "createStyleSheet");
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="bstrHref">optional string bstrHref = </param>
		[CustomMethod]
		[BaseResult]
		[SupportByVersion("MSHTML", 4)]
		public virtual NetOffice.MSHTMLApi.IHTMLStyleSheet createStyleSheet(object bstrHref)
		{
			return InvokerService.InvokeInternal.ExecuteBaseReferenceMethodGet<NetOffice.MSHTMLApi.IHTMLStyleSheet>(this, "createStyleSheet", bstrHref);
		}

		#endregion

		#pragma warning restore
	}
}


