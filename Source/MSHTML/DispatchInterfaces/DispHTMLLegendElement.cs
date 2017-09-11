using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.MSHTMLApi
{
	/// <summary>
	/// DispatchInterface DispHTMLLegendElement 
	/// SupportByVersion MSHTML, 4
	/// </summary>
	[SupportByVersion("MSHTML", 4)]
	[EntityType(EntityType.IsDispatchInterface), BaseType]
 	public class DispHTMLLegendElement : COMObject
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
                    _type = typeof(DispHTMLLegendElement);
                return _type;
            }
        }
        
        #endregion
        
		#region Ctor

		/// <param name="factory">current used factory core</param>
		/// <param name="parentObject">object there has created the proxy</param>
		/// <param name="proxyShare">proxy share instead if com proxy</param>
		public DispHTMLLegendElement(Core factory, ICOMObject parentObject, COMProxyShare proxyShare) : base(factory, parentObject, proxyShare)
		{
		}

		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
		public DispHTMLLegendElement(Core factory, ICOMObject parentObject, object comProxy) : base(factory, parentObject, comProxy)
		{
			
		}

        ///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public DispHTMLLegendElement(ICOMObject parentObject, object comProxy) : base(parentObject, comProxy)
		{
		}
		
		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public DispHTMLLegendElement(Core factory, ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(factory, parentObject, comProxy, comProxyType)
		{

		}

		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public DispHTMLLegendElement(ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(parentObject, comProxy, comProxyType)
		{
		}
		
		///<param name="replacedObject">object to replaced. replacedObject are not usable after this action</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public DispHTMLLegendElement(ICOMObject replacedObject) : base(replacedObject)
		{
		}
		
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public DispHTMLLegendElement() : base()
		{
		}
		
		/// <param name="progId">registered progID</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public DispHTMLLegendElement(string progId) : base(progId)
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
		/// Get
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public string scopeName
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "scopeName");
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public object onlosecapture
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "onlosecapture");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "onlosecapture", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public object onscroll
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "onscroll");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "onscroll", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public object ondrag
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "ondrag");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "ondrag", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public object ondragend
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "ondragend");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "ondragend", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public object ondragenter
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "ondragenter");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "ondragenter", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public object ondragover
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "ondragover");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "ondragover", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public object ondragleave
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "ondragleave");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "ondragleave", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public object ondrop
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "ondrop");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "ondrop", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public object onbeforecut
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "onbeforecut");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "onbeforecut", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public object oncut
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "oncut");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "oncut", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public object onbeforecopy
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "onbeforecopy");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "onbeforecopy", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public object oncopy
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "oncopy");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "oncopy", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public object onbeforepaste
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "onbeforepaste");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "onbeforepaste", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public object onpaste
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "onpaste");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "onpaste", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		[BaseResult]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public NetOffice.MSHTMLApi.IHTMLCurrentStyle currentStyle
		{
			get
			{
				return Factory.ExecuteBaseReferencePropertyGet<NetOffice.MSHTMLApi.IHTMLCurrentStyle>(this, "currentStyle");
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public object onpropertychange
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "onpropertychange");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "onpropertychange", value);
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
		public object readyState
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "readyState");
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public object onreadystatechange
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "onreadystatechange");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "onreadystatechange", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public object onrowsdelete
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "onrowsdelete");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "onrowsdelete", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public object onrowsinserted
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "onrowsinserted");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "onrowsinserted", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public object oncellchange
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "oncellchange");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "oncellchange", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public string dir
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "dir");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "dir", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public Int32 scrollHeight
		{
			get
			{
				return Factory.ExecuteInt32PropertyGet(this, "scrollHeight");
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public Int32 scrollWidth
		{
			get
			{
				return Factory.ExecuteInt32PropertyGet(this, "scrollWidth");
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public Int32 scrollTop
		{
			get
			{
				return Factory.ExecuteInt32PropertyGet(this, "scrollTop");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "scrollTop", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public Int32 scrollLeft
		{
			get
			{
				return Factory.ExecuteInt32PropertyGet(this, "scrollLeft");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "scrollLeft", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public object oncontextmenu
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "oncontextmenu");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "oncontextmenu", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public bool canHaveChildren
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "canHaveChildren");
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		[BaseResult]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public NetOffice.MSHTMLApi.IHTMLStyle runtimeStyle
		{
			get
			{
				return Factory.ExecuteBaseReferencePropertyGet<NetOffice.MSHTMLApi.IHTMLStyle>(this, "runtimeStyle");
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		[SupportByVersion("MSHTML", 4), ProxyResult]
		public object behaviorUrns
		{
			get
			{
				return Factory.ExecuteReferencePropertyGet(this, "behaviorUrns");
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public string tagUrn
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "tagUrn");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "tagUrn", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public object onbeforeeditfocus
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "onbeforeeditfocus");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "onbeforeeditfocus", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Int32 readyStateValue
		{
			get
			{
				return Factory.ExecuteInt32PropertyGet(this, "readyStateValue");
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public bool isMultiLine
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "isMultiLine");
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public bool canHaveHTML
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "canHaveHTML");
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public object onlayoutcomplete
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "onlayoutcomplete");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "onlayoutcomplete", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public object onpage
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "onpage");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "onpage", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public bool inflateBlock
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "inflateBlock");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "inflateBlock", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public object onbeforedeactivate
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "onbeforedeactivate");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "onbeforedeactivate", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public string contentEditable
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "contentEditable");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "contentEditable", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public bool isContentEditable
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "isContentEditable");
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public bool hideFocus
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "hideFocus");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "hideFocus", value);
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
		public bool isDisabled
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "isDisabled");
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public object onmove
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "onmove");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "onmove", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public object oncontrolselect
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "oncontrolselect");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "oncontrolselect", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public object onresizestart
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "onresizestart");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "onresizestart", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public object onresizeend
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "onresizeend");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "onresizeend", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public object onmovestart
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "onmovestart");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "onmovestart", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public object onmoveend
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "onmoveend");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "onmoveend", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public object onmouseenter
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "onmouseenter");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "onmouseenter", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public object onmouseleave
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "onmouseleave");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "onmouseleave", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public object onactivate
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "onactivate");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "onactivate", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public object ondeactivate
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "ondeactivate");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "ondeactivate", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Int32 glyphMode
		{
			get
			{
				return Factory.ExecuteInt32PropertyGet(this, "glyphMode");
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public object onmousewheel
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "onmousewheel");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "onmousewheel", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public object onbeforeactivate
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "onbeforeactivate");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "onbeforeactivate", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public object onfocusin
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "onfocusin");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "onfocusin", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public object onfocusout
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "onfocusout");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "onfocusout", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Int32 uniqueNumber
		{
			get
			{
				return Factory.ExecuteInt32PropertyGet(this, "uniqueNumber");
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public string uniqueID
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "uniqueID");
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public Int32 nodeType
		{
			get
			{
				return Factory.ExecuteInt32PropertyGet(this, "nodeType");
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		[BaseResult]
		public NetOffice.MSHTMLApi.IHTMLDOMNode parentNode
		{
			get
			{
				return Factory.ExecuteBaseReferencePropertyGet<NetOffice.MSHTMLApi.IHTMLDOMNode>(this, "parentNode");
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		[SupportByVersion("MSHTML", 4), ProxyResult]
		public object childNodes
		{
			get
			{
				return Factory.ExecuteReferencePropertyGet(this, "childNodes");
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		[SupportByVersion("MSHTML", 4), ProxyResult]
		public object attributes
		{
			get
			{
				return Factory.ExecuteReferencePropertyGet(this, "attributes");
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public string nodeName
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "nodeName");
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public object nodeValue
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "nodeValue");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "nodeValue", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		[BaseResult]
		public NetOffice.MSHTMLApi.IHTMLDOMNode firstChild
		{
			get
			{
				return Factory.ExecuteBaseReferencePropertyGet<NetOffice.MSHTMLApi.IHTMLDOMNode>(this, "firstChild");
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		[BaseResult]
		public NetOffice.MSHTMLApi.IHTMLDOMNode lastChild
		{
			get
			{
				return Factory.ExecuteBaseReferencePropertyGet<NetOffice.MSHTMLApi.IHTMLDOMNode>(this, "lastChild");
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		[BaseResult]
		public NetOffice.MSHTMLApi.IHTMLDOMNode previousSibling
		{
			get
			{
				return Factory.ExecuteBaseReferencePropertyGet<NetOffice.MSHTMLApi.IHTMLDOMNode>(this, "previousSibling");
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		[BaseResult]
		public NetOffice.MSHTMLApi.IHTMLDOMNode nextSibling
		{
			get
			{
				return Factory.ExecuteBaseReferencePropertyGet<NetOffice.MSHTMLApi.IHTMLDOMNode>(this, "nextSibling");
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		[SupportByVersion("MSHTML", 4), ProxyResult]
		public object ownerDocument
		{
			get
			{
				return Factory.ExecuteReferencePropertyGet(this, "ownerDocument");
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

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public string role
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "role");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "role", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public string ariaBusy
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "ariaBusy");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "ariaBusy", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public string ariaChecked
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "ariaChecked");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "ariaChecked", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public string ariaDisabled
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "ariaDisabled");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "ariaDisabled", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public string ariaExpanded
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "ariaExpanded");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "ariaExpanded", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public string ariaHaspopup
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "ariaHaspopup");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "ariaHaspopup", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public string ariaHidden
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "ariaHidden");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "ariaHidden", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public string ariaInvalid
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "ariaInvalid");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "ariaInvalid", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public string ariaMultiselectable
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "ariaMultiselectable");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "ariaMultiselectable", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public string ariaPressed
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "ariaPressed");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "ariaPressed", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public string ariaReadonly
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "ariaReadonly");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "ariaReadonly", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public string ariaRequired
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "ariaRequired");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "ariaRequired", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public string ariaSecret
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "ariaSecret");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "ariaSecret", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public string ariaSelected
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "ariaSelected");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "ariaSelected", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public NetOffice.MSHTMLApi.IHTMLAttributeCollection3 ie8_attributes
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.MSHTMLApi.IHTMLAttributeCollection3>(this, "ie8_attributes", NetOffice.MSHTMLApi.IHTMLAttributeCollection3.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public string ariaValuenow
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "ariaValuenow");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "ariaValuenow", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public Int16 ariaPosinset
		{
			get
			{
				return Factory.ExecuteInt16PropertyGet(this, "ariaPosinset");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "ariaPosinset", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public Int16 ariaSetsize
		{
			get
			{
				return Factory.ExecuteInt16PropertyGet(this, "ariaSetsize");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "ariaSetsize", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public Int16 ariaLevel
		{
			get
			{
				return Factory.ExecuteInt16PropertyGet(this, "ariaLevel");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "ariaLevel", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public string ariaValuemin
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "ariaValuemin");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "ariaValuemin", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public string ariaValuemax
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "ariaValuemax");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "ariaValuemax", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public string ariaControls
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "ariaControls");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "ariaControls", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public string ariaDescribedby
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "ariaDescribedby");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "ariaDescribedby", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public string ariaFlowto
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "ariaFlowto");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "ariaFlowto", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public string ariaLabelledby
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "ariaLabelledby");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "ariaLabelledby", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public string ariaActivedescendant
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "ariaActivedescendant");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "ariaActivedescendant", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public string ariaOwns
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "ariaOwns");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "ariaOwns", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public string ariaLive
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "ariaLive");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "ariaLive", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public string ariaRelevant
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "ariaRelevant");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "ariaRelevant", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		[SupportByVersion("MSHTML", 4), ProxyResult]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public object constructor
		{
			get
			{
				return Factory.ExecuteReferencePropertyGet(this, "constructor");
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public string align
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "align");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "align", value);
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
		/// <param name="containerCapture">optional bool containerCapture = true</param>
		[SupportByVersion("MSHTML", 4)]
		public void setCapture(object containerCapture)
		{
			 Factory.ExecuteMethod(this, "setCapture", containerCapture);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[CustomMethod]
		[SupportByVersion("MSHTML", 4)]
		public void setCapture()
		{
			 Factory.ExecuteMethod(this, "setCapture");
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public void releaseCapture()
		{
			 Factory.ExecuteMethod(this, "releaseCapture");
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="x">Int32 x</param>
		/// <param name="y">Int32 y</param>
		[SupportByVersion("MSHTML", 4)]
		public string componentFromPoint(Int32 x, Int32 y)
		{
			return Factory.ExecuteStringMethodGet(this, "componentFromPoint", x, y);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="component">optional object component</param>
		[SupportByVersion("MSHTML", 4)]
		public void doScroll(object component)
		{
			 Factory.ExecuteMethod(this, "doScroll", component);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[CustomMethod]
		[SupportByVersion("MSHTML", 4)]
		public void doScroll()
		{
			 Factory.ExecuteMethod(this, "doScroll");
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public NetOffice.MSHTMLApi.IHTMLRectCollection getClientRects()
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.MSHTMLApi.IHTMLRectCollection>(this, "getClientRects", NetOffice.MSHTMLApi.IHTMLRectCollection.LateBindingApiWrapperType);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public NetOffice.MSHTMLApi.IHTMLRect getBoundingClientRect()
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.MSHTMLApi.IHTMLRect>(this, "getBoundingClientRect", NetOffice.MSHTMLApi.IHTMLRect.LateBindingApiWrapperType);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="propname">string propname</param>
		/// <param name="expression">string expression</param>
		/// <param name="language">optional string language = </param>
		[SupportByVersion("MSHTML", 4)]
		public void setExpression(string propname, string expression, object language)
		{
			 Factory.ExecuteMethod(this, "setExpression", propname, expression, language);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="propname">string propname</param>
		/// <param name="expression">string expression</param>
		[CustomMethod]
		[SupportByVersion("MSHTML", 4)]
		public void setExpression(string propname, string expression)
		{
			 Factory.ExecuteMethod(this, "setExpression", propname, expression);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="propname">string propname</param>
		[SupportByVersion("MSHTML", 4)]
		public object getExpression(string propname)
		{
			return Factory.ExecuteVariantMethodGet(this, "getExpression", propname);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="propname">string propname</param>
		[SupportByVersion("MSHTML", 4)]
		public bool removeExpression(string propname)
		{
			return Factory.ExecuteBoolMethodGet(this, "removeExpression", propname);
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
		/// <param name="_event">string event</param>
		/// <param name="pdisp">object pdisp</param>
		[SupportByVersion("MSHTML", 4)]
		public bool attachEvent(string _event, object pdisp)
		{
			return Factory.ExecuteBoolMethodGet(this, "attachEvent", _event, pdisp);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="_event">string event</param>
		/// <param name="pdisp">object pdisp</param>
		[SupportByVersion("MSHTML", 4)]
		public void detachEvent(string _event, object pdisp)
		{
			 Factory.ExecuteMethod(this, "detachEvent", _event, pdisp);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public object createControlRange()
		{
			return Factory.ExecuteVariantMethodGet(this, "createControlRange");
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public void clearAttributes()
		{
			 Factory.ExecuteMethod(this, "clearAttributes");
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="where">string where</param>
		/// <param name="insertedElement">NetOffice.MSHTMLApi.IHTMLElement insertedElement</param>
		[SupportByVersion("MSHTML", 4)]
		[BaseResult]
		public NetOffice.MSHTMLApi.IHTMLElement insertAdjacentElement(string where, NetOffice.MSHTMLApi.IHTMLElement insertedElement)
		{
			return Factory.ExecuteBaseReferenceMethodGet<NetOffice.MSHTMLApi.IHTMLElement>(this, "insertAdjacentElement", where, insertedElement);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="apply">NetOffice.MSHTMLApi.IHTMLElement apply</param>
		/// <param name="where">string where</param>
		[SupportByVersion("MSHTML", 4)]
		[BaseResult]
		public NetOffice.MSHTMLApi.IHTMLElement applyElement(NetOffice.MSHTMLApi.IHTMLElement apply, string where)
		{
			return Factory.ExecuteBaseReferenceMethodGet<NetOffice.MSHTMLApi.IHTMLElement>(this, "applyElement", apply, where);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="where">string where</param>
		[SupportByVersion("MSHTML", 4)]
		public string getAdjacentText(string where)
		{
			return Factory.ExecuteStringMethodGet(this, "getAdjacentText", where);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="where">string where</param>
		/// <param name="newText">string newText</param>
		[SupportByVersion("MSHTML", 4)]
		public string replaceAdjacentText(string where, string newText)
		{
			return Factory.ExecuteStringMethodGet(this, "replaceAdjacentText", where, newText);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="bstrUrl">string bstrUrl</param>
		/// <param name="pvarFactory">optional object pvarFactory</param>
		[SupportByVersion("MSHTML", 4)]
		public Int32 addBehavior(string bstrUrl, object pvarFactory)
		{
			return Factory.ExecuteInt32MethodGet(this, "addBehavior", bstrUrl, pvarFactory);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="bstrUrl">string bstrUrl</param>
		[CustomMethod]
		[SupportByVersion("MSHTML", 4)]
		public Int32 addBehavior(string bstrUrl)
		{
			return Factory.ExecuteInt32MethodGet(this, "addBehavior", bstrUrl);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="cookie">Int32 cookie</param>
		[SupportByVersion("MSHTML", 4)]
		public bool removeBehavior(Int32 cookie)
		{
			return Factory.ExecuteBoolMethodGet(this, "removeBehavior", cookie);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="v">string v</param>
		[SupportByVersion("MSHTML", 4)]
		[BaseResult]
		public NetOffice.MSHTMLApi.IHTMLElementCollection getElementsByTagName(string v)
		{
			return Factory.ExecuteBaseReferenceMethodGet<NetOffice.MSHTMLApi.IHTMLElementCollection>(this, "getElementsByTagName", v);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="mergeThis">NetOffice.MSHTMLApi.IHTMLElement mergeThis</param>
		/// <param name="pvarFlags">optional object pvarFlags</param>
		[SupportByVersion("MSHTML", 4)]
		public void mergeAttributes(NetOffice.MSHTMLApi.IHTMLElement mergeThis, object pvarFlags)
		{
			 Factory.ExecuteMethod(this, "mergeAttributes", mergeThis, pvarFlags);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="mergeThis">NetOffice.MSHTMLApi.IHTMLElement mergeThis</param>
		[CustomMethod]
		[SupportByVersion("MSHTML", 4)]
		public void mergeAttributes(NetOffice.MSHTMLApi.IHTMLElement mergeThis)
		{
			 Factory.ExecuteMethod(this, "mergeAttributes", mergeThis);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public void setActive()
		{
			 Factory.ExecuteMethod(this, "setActive");
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="bstrEventName">string bstrEventName</param>
		/// <param name="pvarEventObject">optional object pvarEventObject</param>
		[SupportByVersion("MSHTML", 4)]
		public bool FireEvent(string bstrEventName, object pvarEventObject)
		{
			return Factory.ExecuteBoolMethodGet(this, "FireEvent", bstrEventName, pvarEventObject);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="bstrEventName">string bstrEventName</param>
		[CustomMethod]
		[SupportByVersion("MSHTML", 4)]
		public bool FireEvent(string bstrEventName)
		{
			return Factory.ExecuteBoolMethodGet(this, "FireEvent", bstrEventName);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public bool dragDrop()
		{
			return Factory.ExecuteBoolMethodGet(this, "dragDrop");
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public void normalize()
		{
			 Factory.ExecuteMethod(this, "normalize");
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="bstrName">string bstrName</param>
		[SupportByVersion("MSHTML", 4)]
		[BaseResult]
		public NetOffice.MSHTMLApi.IHTMLDOMAttribute getAttributeNode(string bstrName)
		{
			return Factory.ExecuteBaseReferenceMethodGet<NetOffice.MSHTMLApi.IHTMLDOMAttribute>(this, "getAttributeNode", bstrName);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="pattr">NetOffice.MSHTMLApi.IHTMLDOMAttribute pattr</param>
		[SupportByVersion("MSHTML", 4)]
		[BaseResult]
		public NetOffice.MSHTMLApi.IHTMLDOMAttribute setAttributeNode(NetOffice.MSHTMLApi.IHTMLDOMAttribute pattr)
		{
			return Factory.ExecuteBaseReferenceMethodGet<NetOffice.MSHTMLApi.IHTMLDOMAttribute>(this, "setAttributeNode", pattr);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="pattr">NetOffice.MSHTMLApi.IHTMLDOMAttribute pattr</param>
		[SupportByVersion("MSHTML", 4)]
		[BaseResult]
		public NetOffice.MSHTMLApi.IHTMLDOMAttribute removeAttributeNode(NetOffice.MSHTMLApi.IHTMLDOMAttribute pattr)
		{
			return Factory.ExecuteBaseReferenceMethodGet<NetOffice.MSHTMLApi.IHTMLDOMAttribute>(this, "removeAttributeNode", pattr);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public bool hasChildNodes()
		{
			return Factory.ExecuteBoolMethodGet(this, "hasChildNodes");
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="newChild">NetOffice.MSHTMLApi.IHTMLDOMNode newChild</param>
		/// <param name="refChild">optional object refChild</param>
		[SupportByVersion("MSHTML", 4)]
		[BaseResult]
		public NetOffice.MSHTMLApi.IHTMLDOMNode insertBefore(NetOffice.MSHTMLApi.IHTMLDOMNode newChild, object refChild)
		{
			return Factory.ExecuteBaseReferenceMethodGet<NetOffice.MSHTMLApi.IHTMLDOMNode>(this, "insertBefore", newChild, refChild);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="newChild">NetOffice.MSHTMLApi.IHTMLDOMNode newChild</param>
		[CustomMethod]
		[BaseResult]
		[SupportByVersion("MSHTML", 4)]
		public NetOffice.MSHTMLApi.IHTMLDOMNode insertBefore(NetOffice.MSHTMLApi.IHTMLDOMNode newChild)
		{
			return Factory.ExecuteBaseReferenceMethodGet<NetOffice.MSHTMLApi.IHTMLDOMNode>(this, "insertBefore", newChild);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="oldChild">NetOffice.MSHTMLApi.IHTMLDOMNode oldChild</param>
		[SupportByVersion("MSHTML", 4)]
		[BaseResult]
		public NetOffice.MSHTMLApi.IHTMLDOMNode removeChild(NetOffice.MSHTMLApi.IHTMLDOMNode oldChild)
		{
			return Factory.ExecuteBaseReferenceMethodGet<NetOffice.MSHTMLApi.IHTMLDOMNode>(this, "removeChild", oldChild);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="newChild">NetOffice.MSHTMLApi.IHTMLDOMNode newChild</param>
		/// <param name="oldChild">NetOffice.MSHTMLApi.IHTMLDOMNode oldChild</param>
		[SupportByVersion("MSHTML", 4)]
		[BaseResult]
		public NetOffice.MSHTMLApi.IHTMLDOMNode replaceChild(NetOffice.MSHTMLApi.IHTMLDOMNode newChild, NetOffice.MSHTMLApi.IHTMLDOMNode oldChild)
		{
			return Factory.ExecuteBaseReferenceMethodGet<NetOffice.MSHTMLApi.IHTMLDOMNode>(this, "replaceChild", newChild, oldChild);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="fDeep">bool fDeep</param>
		[SupportByVersion("MSHTML", 4)]
		[BaseResult]
		public NetOffice.MSHTMLApi.IHTMLDOMNode cloneNode(bool fDeep)
		{
			return Factory.ExecuteBaseReferenceMethodGet<NetOffice.MSHTMLApi.IHTMLDOMNode>(this, "cloneNode", fDeep);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="fDeep">optional bool fDeep = false</param>
		[SupportByVersion("MSHTML", 4)]
		[BaseResult]
		public NetOffice.MSHTMLApi.IHTMLDOMNode removeNode(object fDeep)
		{
			return Factory.ExecuteBaseReferenceMethodGet<NetOffice.MSHTMLApi.IHTMLDOMNode>(this, "removeNode", fDeep);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[CustomMethod]
		[BaseResult]
		[SupportByVersion("MSHTML", 4)]
		public NetOffice.MSHTMLApi.IHTMLDOMNode removeNode()
		{
			return Factory.ExecuteBaseReferenceMethodGet<NetOffice.MSHTMLApi.IHTMLDOMNode>(this, "removeNode");
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="otherNode">NetOffice.MSHTMLApi.IHTMLDOMNode otherNode</param>
		[SupportByVersion("MSHTML", 4)]
		[BaseResult]
		public NetOffice.MSHTMLApi.IHTMLDOMNode swapNode(NetOffice.MSHTMLApi.IHTMLDOMNode otherNode)
		{
			return Factory.ExecuteBaseReferenceMethodGet<NetOffice.MSHTMLApi.IHTMLDOMNode>(this, "swapNode", otherNode);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="replacement">NetOffice.MSHTMLApi.IHTMLDOMNode replacement</param>
		[SupportByVersion("MSHTML", 4)]
		[BaseResult]
		public NetOffice.MSHTMLApi.IHTMLDOMNode replaceNode(NetOffice.MSHTMLApi.IHTMLDOMNode replacement)
		{
			return Factory.ExecuteBaseReferenceMethodGet<NetOffice.MSHTMLApi.IHTMLDOMNode>(this, "replaceNode", replacement);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="newChild">NetOffice.MSHTMLApi.IHTMLDOMNode newChild</param>
		[SupportByVersion("MSHTML", 4)]
		[BaseResult]
		public NetOffice.MSHTMLApi.IHTMLDOMNode appendChild(NetOffice.MSHTMLApi.IHTMLDOMNode newChild)
		{
			return Factory.ExecuteBaseReferenceMethodGet<NetOffice.MSHTMLApi.IHTMLDOMNode>(this, "appendChild", newChild);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="bstrName">string bstrName</param>
		[SupportByVersion("MSHTML", 4)]
		[BaseResult]
		public NetOffice.MSHTMLApi.IHTMLDOMAttribute2 ie8_getAttributeNode(string bstrName)
		{
			return Factory.ExecuteBaseReferenceMethodGet<NetOffice.MSHTMLApi.IHTMLDOMAttribute2>(this, "ie8_getAttributeNode", bstrName);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="pattr">NetOffice.MSHTMLApi.IHTMLDOMAttribute2 pattr</param>
		[SupportByVersion("MSHTML", 4)]
		[BaseResult]
		public NetOffice.MSHTMLApi.IHTMLDOMAttribute2 ie8_setAttributeNode(NetOffice.MSHTMLApi.IHTMLDOMAttribute2 pattr)
		{
			return Factory.ExecuteBaseReferenceMethodGet<NetOffice.MSHTMLApi.IHTMLDOMAttribute2>(this, "ie8_setAttributeNode", pattr);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="pattr">NetOffice.MSHTMLApi.IHTMLDOMAttribute2 pattr</param>
		[SupportByVersion("MSHTML", 4)]
		[BaseResult]
		public NetOffice.MSHTMLApi.IHTMLDOMAttribute2 ie8_removeAttributeNode(NetOffice.MSHTMLApi.IHTMLDOMAttribute2 pattr)
		{
			return Factory.ExecuteBaseReferenceMethodGet<NetOffice.MSHTMLApi.IHTMLDOMAttribute2>(this, "ie8_removeAttributeNode", pattr);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="name">string name</param>
		[SupportByVersion("MSHTML", 4)]
		public bool hasAttribute(string name)
		{
			return Factory.ExecuteBoolMethodGet(this, "hasAttribute", name);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="strAttributeName">string strAttributeName</param>
		[SupportByVersion("MSHTML", 4)]
		public object ie8_getAttribute(string strAttributeName)
		{
			return Factory.ExecuteVariantMethodGet(this, "ie8_getAttribute", strAttributeName);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="strAttributeName">string strAttributeName</param>
		/// <param name="attributeValue">object attributeValue</param>
		[SupportByVersion("MSHTML", 4)]
		public void ie8_setAttribute(string strAttributeName, object attributeValue)
		{
			 Factory.ExecuteMethod(this, "ie8_setAttribute", strAttributeName, attributeValue);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="strAttributeName">string strAttributeName</param>
		[SupportByVersion("MSHTML", 4)]
		public bool ie8_removeAttribute(string strAttributeName)
		{
			return Factory.ExecuteBoolMethodGet(this, "ie8_removeAttribute", strAttributeName);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public bool hasAttributes()
		{
			return Factory.ExecuteBoolMethodGet(this, "hasAttributes");
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="v">string v</param>
		[SupportByVersion("MSHTML", 4)]
		[BaseResult]
		public NetOffice.MSHTMLApi.IHTMLElement querySelector(string v)
		{
			return Factory.ExecuteBaseReferenceMethodGet<NetOffice.MSHTMLApi.IHTMLElement>(this, "querySelector", v);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="v">string v</param>
		[SupportByVersion("MSHTML", 4)]
		public NetOffice.MSHTMLApi.IHTMLDOMChildrenCollection querySelectorAll(string v)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.MSHTMLApi.IHTMLDOMChildrenCollection>(this, "querySelectorAll", NetOffice.MSHTMLApi.IHTMLDOMChildrenCollection.LateBindingApiWrapperType, v);
		}

		#endregion

		#pragma warning restore
	}
}
