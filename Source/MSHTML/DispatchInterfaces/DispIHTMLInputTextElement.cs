using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.MSHTMLApi
{
	/// <summary>
	/// DispatchInterface DispIHTMLInputTextElement 
	/// SupportByVersion MSHTML, 4
	/// </summary>
	[SupportByVersion("MSHTML", 4)]
	[EntityType(EntityType.IsDispatchInterface), BaseType]
 	public class DispIHTMLInputTextElement : COMObject
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
                    _type = typeof(DispIHTMLInputTextElement);
                return _type;
            }
        }
        
        #endregion
        
		#region Ctor

		/// <param name="factory">current used factory core</param>
		/// <param name="parentObject">object there has created the proxy</param>
		/// <param name="proxyShare">proxy share instead if com proxy</param>
		public DispIHTMLInputTextElement(Core factory, ICOMObject parentObject, COMProxyShare proxyShare) : base(factory, parentObject, proxyShare)
		{
		}

		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
		public DispIHTMLInputTextElement(Core factory, ICOMObject parentObject, object comProxy) : base(factory, parentObject, comProxy)
		{
			
		}

        ///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public DispIHTMLInputTextElement(ICOMObject parentObject, object comProxy) : base(parentObject, comProxy)
		{
		}
		
		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public DispIHTMLInputTextElement(Core factory, ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(factory, parentObject, comProxy, comProxyType)
		{

		}

		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public DispIHTMLInputTextElement(ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(parentObject, comProxy, comProxyType)
		{
		}
		
		///<param name="replacedObject">object to replaced. replacedObject are not usable after this action</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public DispIHTMLInputTextElement(ICOMObject replacedObject) : base(replacedObject)
		{
		}
		
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public DispIHTMLInputTextElement() : base()
		{
		}
		
		/// <param name="progId">registered progID</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public DispIHTMLInputTextElement(string progId) : base(progId)
		{
		}
		
		#endregion
		
		#region Properties

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public string className
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "className");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "className", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public string id
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "id");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "id", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public string tagName
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "tagName");
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		[BaseResult]
		public NetOffice.MSHTMLApi.IHTMLElement parentElement
		{
			get
			{
				return Factory.ExecuteBaseReferencePropertyGet<NetOffice.MSHTMLApi.IHTMLElement>(this, "parentElement");
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		[BaseResult]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public NetOffice.MSHTMLApi.IHTMLStyle style
		{
			get
			{
				return Factory.ExecuteBaseReferencePropertyGet<NetOffice.MSHTMLApi.IHTMLStyle>(this, "style");
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public object onhelp
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "onhelp");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "onhelp", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public object onclick
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "onclick");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "onclick", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public object ondblclick
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "ondblclick");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "ondblclick", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public object onkeydown
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "onkeydown");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "onkeydown", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public object onkeyup
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "onkeyup");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "onkeyup", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public object onkeypress
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "onkeypress");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "onkeypress", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public object onmouseout
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "onmouseout");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "onmouseout", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public object onmouseover
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "onmouseover");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "onmouseover", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public object onmousemove
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "onmousemove");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "onmousemove", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public object onmousedown
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "onmousedown");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "onmousedown", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public object onmouseup
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "onmouseup");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "onmouseup", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		[SupportByVersion("MSHTML", 4), ProxyResult]
		public object document
		{
			get
			{
				return Factory.ExecuteReferencePropertyGet(this, "document");
			}
		}

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
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public string language
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "language");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "language", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public object onselectstart
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "onselectstart");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "onselectstart", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public Int32 sourceIndex
		{
			get
			{
				return Factory.ExecuteInt32PropertyGet(this, "sourceIndex");
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public object recordNumber
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "recordNumber");
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public string lang
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "lang");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "lang", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public Int32 offsetLeft
		{
			get
			{
				return Factory.ExecuteInt32PropertyGet(this, "offsetLeft");
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public Int32 offsetTop
		{
			get
			{
				return Factory.ExecuteInt32PropertyGet(this, "offsetTop");
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public Int32 offsetWidth
		{
			get
			{
				return Factory.ExecuteInt32PropertyGet(this, "offsetWidth");
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public Int32 offsetHeight
		{
			get
			{
				return Factory.ExecuteInt32PropertyGet(this, "offsetHeight");
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		[BaseResult]
		public NetOffice.MSHTMLApi.IHTMLElement offsetParent
		{
			get
			{
				return Factory.ExecuteBaseReferencePropertyGet<NetOffice.MSHTMLApi.IHTMLElement>(this, "offsetParent");
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public string innerHTML
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "innerHTML");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "innerHTML", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public string innerText
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "innerText");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "innerText", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public string outerHTML
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "outerHTML");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "outerHTML", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public string outerText
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "outerText");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "outerText", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		[BaseResult]
		public NetOffice.MSHTMLApi.IHTMLElement parentTextEdit
		{
			get
			{
				return Factory.ExecuteBaseReferencePropertyGet<NetOffice.MSHTMLApi.IHTMLElement>(this, "parentTextEdit");
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public bool isTextEdit
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "isTextEdit");
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public NetOffice.MSHTMLApi.IHTMLFiltersCollection filters
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.MSHTMLApi.IHTMLFiltersCollection>(this, "filters", NetOffice.MSHTMLApi.IHTMLFiltersCollection.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public object ondragstart
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "ondragstart");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "ondragstart", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public object onbeforeupdate
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "onbeforeupdate");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "onbeforeupdate", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public object onafterupdate
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "onafterupdate");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "onafterupdate", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public object onerrorupdate
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "onerrorupdate");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "onerrorupdate", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public object onrowexit
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "onrowexit");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "onrowexit", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public object onrowenter
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "onrowenter");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "onrowenter", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public object ondatasetchanged
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "ondatasetchanged");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "ondatasetchanged", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public object ondataavailable
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "ondataavailable");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "ondataavailable", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public object ondatasetcomplete
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "ondatasetcomplete");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "ondatasetcomplete", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public object onfilterchange
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "onfilterchange");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "onfilterchange", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		[SupportByVersion("MSHTML", 4), ProxyResult]
		public object children
		{
			get
			{
				return Factory.ExecuteReferencePropertyGet(this, "children");
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		[SupportByVersion("MSHTML", 4), ProxyResult]
		public object all
		{
			get
			{
				return Factory.ExecuteReferencePropertyGet(this, "all");
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public Int16 tabIndex
		{
			get
			{
				return Factory.ExecuteInt16PropertyGet(this, "tabIndex");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "tabIndex", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public string accessKey
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "accessKey");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "accessKey", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public object onblur
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "onblur");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "onblur", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public object onfocus
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "onfocus");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "onfocus", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public object onresize
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "onresize");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "onresize", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public Int32 clientHeight
		{
			get
			{
				return Factory.ExecuteInt32PropertyGet(this, "clientHeight");
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public Int32 clientWidth
		{
			get
			{
				return Factory.ExecuteInt32PropertyGet(this, "clientWidth");
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public Int32 clientTop
		{
			get
			{
				return Factory.ExecuteInt32PropertyGet(this, "clientTop");
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public Int32 clientLeft
		{
			get
			{
				return Factory.ExecuteInt32PropertyGet(this, "clientLeft");
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
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public string value
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "value");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "value", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public string name
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "name");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "name", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public object status
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "status");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "status", value);
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
		[BaseResult]
		public NetOffice.MSHTMLApi.IHTMLFormElement form
		{
			get
			{
				return Factory.ExecuteBaseReferencePropertyGet<NetOffice.MSHTMLApi.IHTMLFormElement>(this, "form");
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public string defaultValue
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "defaultValue");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "defaultValue", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public Int32 size
		{
			get
			{
				return Factory.ExecuteInt32PropertyGet(this, "size");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "size", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public Int32 maxLength
		{
			get
			{
				return Factory.ExecuteInt32PropertyGet(this, "maxLength");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "maxLength", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public object onchange
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "onchange");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "onchange", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public object onselect
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "onselect");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "onselect", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public bool readOnly
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "readOnly");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "readOnly", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public string dataFld
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "dataFld");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "dataFld", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public string dataSrc
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "dataSrc");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "dataSrc", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public string dataFormatAs
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "dataFormatAs");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "dataFormatAs", value);
			}
		}

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="strAttributeName">string strAttributeName</param>
		/// <param name="attributeValue">object attributeValue</param>
		/// <param name="lFlags">optional Int32 lFlags = 1</param>
		[SupportByVersion("MSHTML", 4)]
		public void setAttribute(string strAttributeName, object attributeValue, object lFlags)
		{
			 Factory.ExecuteMethod(this, "setAttribute", strAttributeName, attributeValue, lFlags);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="strAttributeName">string strAttributeName</param>
		/// <param name="attributeValue">object attributeValue</param>
		[CustomMethod]
		[SupportByVersion("MSHTML", 4)]
		public void setAttribute(string strAttributeName, object attributeValue)
		{
			 Factory.ExecuteMethod(this, "setAttribute", strAttributeName, attributeValue);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="strAttributeName">string strAttributeName</param>
		/// <param name="lFlags">optional Int32 lFlags = 0</param>
		[SupportByVersion("MSHTML", 4)]
		public object getAttribute(string strAttributeName, object lFlags)
		{
			return Factory.ExecuteVariantMethodGet(this, "getAttribute", strAttributeName, lFlags);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="strAttributeName">string strAttributeName</param>
		[CustomMethod]
		[SupportByVersion("MSHTML", 4)]
		public object getAttribute(string strAttributeName)
		{
			return Factory.ExecuteVariantMethodGet(this, "getAttribute", strAttributeName);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="strAttributeName">string strAttributeName</param>
		/// <param name="lFlags">optional Int32 lFlags = 1</param>
		[SupportByVersion("MSHTML", 4)]
		public bool removeAttribute(string strAttributeName, object lFlags)
		{
			return Factory.ExecuteBoolMethodGet(this, "removeAttribute", strAttributeName, lFlags);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="strAttributeName">string strAttributeName</param>
		[CustomMethod]
		[SupportByVersion("MSHTML", 4)]
		public bool removeAttribute(string strAttributeName)
		{
			return Factory.ExecuteBoolMethodGet(this, "removeAttribute", strAttributeName);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="varargStart">optional object varargStart</param>
		[SupportByVersion("MSHTML", 4)]
		public void scrollIntoView(object varargStart)
		{
			 Factory.ExecuteMethod(this, "scrollIntoView", varargStart);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[CustomMethod]
		[SupportByVersion("MSHTML", 4)]
		public void scrollIntoView()
		{
			 Factory.ExecuteMethod(this, "scrollIntoView");
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="pChild">NetOffice.MSHTMLApi.IHTMLElement pChild</param>
		[SupportByVersion("MSHTML", 4)]
		public bool contains(NetOffice.MSHTMLApi.IHTMLElement pChild)
		{
			return Factory.ExecuteBoolMethodGet(this, "contains", pChild);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="where">string where</param>
		/// <param name="html">string html</param>
		[SupportByVersion("MSHTML", 4)]
		public void insertAdjacentHTML(string where, string html)
		{
			 Factory.ExecuteMethod(this, "insertAdjacentHTML", where, html);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="where">string where</param>
		/// <param name="text">string text</param>
		[SupportByVersion("MSHTML", 4)]
		public void insertAdjacentText(string where, string text)
		{
			 Factory.ExecuteMethod(this, "insertAdjacentText", where, text);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public void click()
		{
			 Factory.ExecuteMethod(this, "click");
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public string toString()
		{
			return Factory.ExecuteStringMethodGet(this, "toString");
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public void focus()
		{
			 Factory.ExecuteMethod(this, "focus");
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public void blur()
		{
			 Factory.ExecuteMethod(this, "blur");
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="pUnk">object pUnk</param>
		[SupportByVersion("MSHTML", 4)]
		public void addFilter(object pUnk)
		{
			 Factory.ExecuteMethod(this, "addFilter", pUnk);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="pUnk">object pUnk</param>
		[SupportByVersion("MSHTML", 4)]
		public void removeFilter(object pUnk)
		{
			 Factory.ExecuteMethod(this, "removeFilter", pUnk);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public void select()
		{
			 Factory.ExecuteMethod(this, "select");
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public NetOffice.MSHTMLApi.IHTMLTxtRange createTextRange()
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.MSHTMLApi.IHTMLTxtRange>(this, "createTextRange", NetOffice.MSHTMLApi.IHTMLTxtRange.LateBindingApiWrapperType);
		}

		#endregion

		#pragma warning restore
	}
}
