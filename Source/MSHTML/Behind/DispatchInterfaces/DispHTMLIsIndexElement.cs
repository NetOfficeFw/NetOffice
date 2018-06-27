using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;
using NetOffice.MSHTMLApi;

namespace NetOffice.MSHTMLApi.Behind
{
	/// <summary>
	/// DispatchInterface DispHTMLIsIndexElement 
	/// SupportByVersion MSHTML, 4
	/// </summary>
	[SupportByVersion("MSHTML", 4)]
	[EntityType(EntityType.IsDispatchInterface), BaseType]
 	public class DispHTMLIsIndexElement : COMObject, NetOffice.MSHTMLApi.DispHTMLIsIndexElement
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
                    _contractType = typeof(NetOffice.MSHTMLApi.DispHTMLIsIndexElement);
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
                    _type = typeof(DispHTMLIsIndexElement);
                return _type;
            }
        }
        
        #endregion
        
		#region Ctor

		/// <summary>
		/// Stub Ctor, not indented to use
		/// </summary>
		public DispHTMLIsIndexElement() : base()
		{

		}

		#endregion
		
		#region Properties

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public virtual string className
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "className");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "className", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public virtual string id
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "id");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "id", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public virtual string tagName
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "tagName");
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		[BaseResult]
		public virtual NetOffice.MSHTMLApi.IHTMLElement parentElement
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteBaseReferencePropertyGet<NetOffice.MSHTMLApi.IHTMLElement>(this, "parentElement");
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		[BaseResult]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public virtual NetOffice.MSHTMLApi.IHTMLStyle style
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteBaseReferencePropertyGet<NetOffice.MSHTMLApi.IHTMLStyle>(this, "style");
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
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		[SupportByVersion("MSHTML", 4), ProxyResult]
		public virtual object document
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteReferencePropertyGet(this, "document");
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
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public virtual string language
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "language");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "language", value);
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
		public virtual Int32 sourceIndex
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteInt32PropertyGet(this, "sourceIndex");
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public virtual object recordNumber
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "recordNumber");
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public virtual string lang
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "lang");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "lang", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public virtual Int32 offsetLeft
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteInt32PropertyGet(this, "offsetLeft");
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public virtual Int32 offsetTop
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteInt32PropertyGet(this, "offsetTop");
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public virtual Int32 offsetWidth
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteInt32PropertyGet(this, "offsetWidth");
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public virtual Int32 offsetHeight
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteInt32PropertyGet(this, "offsetHeight");
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		[BaseResult]
		public virtual NetOffice.MSHTMLApi.IHTMLElement offsetParent
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteBaseReferencePropertyGet<NetOffice.MSHTMLApi.IHTMLElement>(this, "offsetParent");
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public virtual string innerHTML
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "innerHTML");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "innerHTML", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public virtual string innerText
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "innerText");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "innerText", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public virtual string outerHTML
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "outerHTML");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "outerHTML", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public virtual string outerText
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "outerText");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "outerText", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		[BaseResult]
		public virtual NetOffice.MSHTMLApi.IHTMLElement parentTextEdit
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteBaseReferencePropertyGet<NetOffice.MSHTMLApi.IHTMLElement>(this, "parentTextEdit");
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public virtual bool isTextEdit
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "isTextEdit");
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public virtual NetOffice.MSHTMLApi.IHTMLFiltersCollection filters
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.MSHTMLApi.IHTMLFiltersCollection>(this, "filters", typeof(NetOffice.MSHTMLApi.IHTMLFiltersCollection));
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
		public virtual object ondatasetchanged
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "ondatasetchanged");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteVariantPropertySet(this, "ondatasetchanged", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public virtual object ondataavailable
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "ondataavailable");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteVariantPropertySet(this, "ondataavailable", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public virtual object ondatasetcomplete
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "ondatasetcomplete");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteVariantPropertySet(this, "ondatasetcomplete", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public virtual object onfilterchange
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "onfilterchange");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteVariantPropertySet(this, "onfilterchange", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		[SupportByVersion("MSHTML", 4), ProxyResult]
		public virtual object children
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteReferencePropertyGet(this, "children");
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		[SupportByVersion("MSHTML", 4), ProxyResult]
		public virtual object all
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteReferencePropertyGet(this, "all");
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public virtual string scopeName
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "scopeName");
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public virtual object onlosecapture
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "onlosecapture");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteVariantPropertySet(this, "onlosecapture", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public virtual object onscroll
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "onscroll");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteVariantPropertySet(this, "onscroll", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public virtual object ondrag
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "ondrag");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteVariantPropertySet(this, "ondrag", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public virtual object ondragend
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "ondragend");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteVariantPropertySet(this, "ondragend", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public virtual object ondragenter
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "ondragenter");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteVariantPropertySet(this, "ondragenter", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public virtual object ondragover
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "ondragover");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteVariantPropertySet(this, "ondragover", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public virtual object ondragleave
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "ondragleave");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteVariantPropertySet(this, "ondragleave", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public virtual object ondrop
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "ondrop");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteVariantPropertySet(this, "ondrop", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public virtual object onbeforecut
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "onbeforecut");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteVariantPropertySet(this, "onbeforecut", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public virtual object oncut
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "oncut");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteVariantPropertySet(this, "oncut", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public virtual object onbeforecopy
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "onbeforecopy");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteVariantPropertySet(this, "onbeforecopy", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public virtual object oncopy
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "oncopy");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteVariantPropertySet(this, "oncopy", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public virtual object onbeforepaste
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "onbeforepaste");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteVariantPropertySet(this, "onbeforepaste", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public virtual object onpaste
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "onpaste");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteVariantPropertySet(this, "onpaste", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		[BaseResult]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public virtual NetOffice.MSHTMLApi.IHTMLCurrentStyle currentStyle
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteBaseReferencePropertyGet<NetOffice.MSHTMLApi.IHTMLCurrentStyle>(this, "currentStyle");
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public virtual object onpropertychange
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "onpropertychange");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteVariantPropertySet(this, "onpropertychange", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public virtual Int16 tabIndex
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteInt16PropertyGet(this, "tabIndex");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "tabIndex", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public virtual string accessKey
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "accessKey");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "accessKey", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public virtual object onblur
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "onblur");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteVariantPropertySet(this, "onblur", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public virtual object onfocus
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "onfocus");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteVariantPropertySet(this, "onfocus", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public virtual object onresize
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "onresize");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteVariantPropertySet(this, "onresize", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public virtual Int32 clientHeight
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteInt32PropertyGet(this, "clientHeight");
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public virtual Int32 clientWidth
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteInt32PropertyGet(this, "clientWidth");
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public virtual Int32 clientTop
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteInt32PropertyGet(this, "clientTop");
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public virtual Int32 clientLeft
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteInt32PropertyGet(this, "clientLeft");
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public virtual object readyState
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "readyState");
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
		public virtual object onrowsdelete
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "onrowsdelete");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteVariantPropertySet(this, "onrowsdelete", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public virtual object onrowsinserted
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "onrowsinserted");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteVariantPropertySet(this, "onrowsinserted", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public virtual object oncellchange
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "oncellchange");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteVariantPropertySet(this, "oncellchange", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public virtual string dir
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "dir");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "dir", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public virtual Int32 scrollHeight
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteInt32PropertyGet(this, "scrollHeight");
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public virtual Int32 scrollWidth
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteInt32PropertyGet(this, "scrollWidth");
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public virtual Int32 scrollTop
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteInt32PropertyGet(this, "scrollTop");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "scrollTop", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public virtual Int32 scrollLeft
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteInt32PropertyGet(this, "scrollLeft");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "scrollLeft", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public virtual object oncontextmenu
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "oncontextmenu");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteVariantPropertySet(this, "oncontextmenu", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public virtual bool canHaveChildren
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "canHaveChildren");
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		[BaseResult]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public virtual NetOffice.MSHTMLApi.IHTMLStyle runtimeStyle
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteBaseReferencePropertyGet<NetOffice.MSHTMLApi.IHTMLStyle>(this, "runtimeStyle");
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		[SupportByVersion("MSHTML", 4), ProxyResult]
		public virtual object behaviorUrns
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteReferencePropertyGet(this, "behaviorUrns");
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public virtual string tagUrn
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "tagUrn");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "tagUrn", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public virtual object onbeforeeditfocus
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "onbeforeeditfocus");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteVariantPropertySet(this, "onbeforeeditfocus", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public virtual Int32 readyStateValue
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteInt32PropertyGet(this, "readyStateValue");
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public virtual bool isMultiLine
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "isMultiLine");
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public virtual bool canHaveHTML
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "canHaveHTML");
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public virtual object onlayoutcomplete
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "onlayoutcomplete");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteVariantPropertySet(this, "onlayoutcomplete", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public virtual object onpage
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "onpage");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteVariantPropertySet(this, "onpage", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public virtual bool inflateBlock
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "inflateBlock");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "inflateBlock", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public virtual object onbeforedeactivate
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "onbeforedeactivate");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteVariantPropertySet(this, "onbeforedeactivate", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public virtual string contentEditable
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "contentEditable");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "contentEditable", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public virtual bool isContentEditable
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "isContentEditable");
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public virtual bool hideFocus
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "hideFocus");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "hideFocus", value);
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
		public virtual bool isDisabled
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "isDisabled");
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public virtual object onmove
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "onmove");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteVariantPropertySet(this, "onmove", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public virtual object oncontrolselect
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "oncontrolselect");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteVariantPropertySet(this, "oncontrolselect", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public virtual object onresizestart
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "onresizestart");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteVariantPropertySet(this, "onresizestart", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public virtual object onresizeend
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "onresizeend");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteVariantPropertySet(this, "onresizeend", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public virtual object onmovestart
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "onmovestart");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteVariantPropertySet(this, "onmovestart", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public virtual object onmoveend
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "onmoveend");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteVariantPropertySet(this, "onmoveend", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public virtual object onmouseenter
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "onmouseenter");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteVariantPropertySet(this, "onmouseenter", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public virtual object onmouseleave
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "onmouseleave");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteVariantPropertySet(this, "onmouseleave", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public virtual object onactivate
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "onactivate");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteVariantPropertySet(this, "onactivate", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public virtual object ondeactivate
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "ondeactivate");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteVariantPropertySet(this, "ondeactivate", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public virtual Int32 glyphMode
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteInt32PropertyGet(this, "glyphMode");
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public virtual object onmousewheel
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "onmousewheel");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteVariantPropertySet(this, "onmousewheel", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public virtual object onbeforeactivate
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "onbeforeactivate");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteVariantPropertySet(this, "onbeforeactivate", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public virtual object onfocusin
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "onfocusin");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteVariantPropertySet(this, "onfocusin", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public virtual object onfocusout
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "onfocusout");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteVariantPropertySet(this, "onfocusout", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public virtual Int32 uniqueNumber
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteInt32PropertyGet(this, "uniqueNumber");
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public virtual string uniqueID
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "uniqueID");
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public virtual Int32 nodeType
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteInt32PropertyGet(this, "nodeType");
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		[BaseResult]
		public virtual NetOffice.MSHTMLApi.IHTMLDOMNode parentNode
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteBaseReferencePropertyGet<NetOffice.MSHTMLApi.IHTMLDOMNode>(this, "parentNode");
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		[SupportByVersion("MSHTML", 4), ProxyResult]
		public virtual object childNodes
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteReferencePropertyGet(this, "childNodes");
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		[SupportByVersion("MSHTML", 4), ProxyResult]
		public virtual object attributes
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteReferencePropertyGet(this, "attributes");
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public virtual string nodeName
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "nodeName");
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public virtual object nodeValue
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "nodeValue");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteVariantPropertySet(this, "nodeValue", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		[BaseResult]
		public virtual NetOffice.MSHTMLApi.IHTMLDOMNode firstChild
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteBaseReferencePropertyGet<NetOffice.MSHTMLApi.IHTMLDOMNode>(this, "firstChild");
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		[BaseResult]
		public virtual NetOffice.MSHTMLApi.IHTMLDOMNode lastChild
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteBaseReferencePropertyGet<NetOffice.MSHTMLApi.IHTMLDOMNode>(this, "lastChild");
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		[BaseResult]
		public virtual NetOffice.MSHTMLApi.IHTMLDOMNode previousSibling
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteBaseReferencePropertyGet<NetOffice.MSHTMLApi.IHTMLDOMNode>(this, "previousSibling");
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		[BaseResult]
		public virtual NetOffice.MSHTMLApi.IHTMLDOMNode nextSibling
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteBaseReferencePropertyGet<NetOffice.MSHTMLApi.IHTMLDOMNode>(this, "nextSibling");
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		[SupportByVersion("MSHTML", 4), ProxyResult]
		public virtual object ownerDocument
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteReferencePropertyGet(this, "ownerDocument");
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public virtual string role
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "role");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "role", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public virtual string ariaBusy
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "ariaBusy");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "ariaBusy", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public virtual string ariaChecked
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "ariaChecked");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "ariaChecked", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public virtual string ariaDisabled
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "ariaDisabled");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "ariaDisabled", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public virtual string ariaExpanded
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "ariaExpanded");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "ariaExpanded", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public virtual string ariaHaspopup
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "ariaHaspopup");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "ariaHaspopup", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public virtual string ariaHidden
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "ariaHidden");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "ariaHidden", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public virtual string ariaInvalid
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "ariaInvalid");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "ariaInvalid", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public virtual string ariaMultiselectable
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "ariaMultiselectable");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "ariaMultiselectable", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public virtual string ariaPressed
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "ariaPressed");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "ariaPressed", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public virtual string ariaReadonly
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "ariaReadonly");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "ariaReadonly", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public virtual string ariaRequired
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "ariaRequired");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "ariaRequired", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public virtual string ariaSecret
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "ariaSecret");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "ariaSecret", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public virtual string ariaSelected
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "ariaSelected");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "ariaSelected", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public virtual NetOffice.MSHTMLApi.IHTMLAttributeCollection3 ie8_attributes
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.MSHTMLApi.IHTMLAttributeCollection3>(this, "ie8_attributes", typeof(NetOffice.MSHTMLApi.IHTMLAttributeCollection3));
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public virtual string ariaValuenow
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "ariaValuenow");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "ariaValuenow", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public virtual Int16 ariaPosinset
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteInt16PropertyGet(this, "ariaPosinset");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "ariaPosinset", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public virtual Int16 ariaSetsize
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteInt16PropertyGet(this, "ariaSetsize");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "ariaSetsize", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public virtual Int16 ariaLevel
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteInt16PropertyGet(this, "ariaLevel");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "ariaLevel", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public virtual string ariaValuemin
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "ariaValuemin");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "ariaValuemin", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public virtual string ariaValuemax
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "ariaValuemax");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "ariaValuemax", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public virtual string ariaControls
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "ariaControls");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "ariaControls", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public virtual string ariaDescribedby
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "ariaDescribedby");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "ariaDescribedby", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public virtual string ariaFlowto
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "ariaFlowto");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "ariaFlowto", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public virtual string ariaLabelledby
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "ariaLabelledby");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "ariaLabelledby", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public virtual string ariaActivedescendant
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "ariaActivedescendant");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "ariaActivedescendant", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public virtual string ariaOwns
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "ariaOwns");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "ariaOwns", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public virtual string ariaLive
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "ariaLive");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "ariaLive", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public virtual string ariaRelevant
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "ariaRelevant");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "ariaRelevant", value);
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

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public virtual string prompt
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "prompt");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "prompt", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public virtual string action
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "action");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "action", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		[BaseResult]
		public virtual NetOffice.MSHTMLApi.IHTMLFormElement form
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteBaseReferencePropertyGet<NetOffice.MSHTMLApi.IHTMLFormElement>(this, "form");
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
		public virtual void setAttribute(string strAttributeName, object attributeValue, object lFlags)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "setAttribute", strAttributeName, attributeValue, lFlags);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="strAttributeName">string strAttributeName</param>
		/// <param name="attributeValue">object attributeValue</param>
		[CustomMethod]
		[SupportByVersion("MSHTML", 4)]
		public virtual void setAttribute(string strAttributeName, object attributeValue)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "setAttribute", strAttributeName, attributeValue);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="strAttributeName">string strAttributeName</param>
		/// <param name="lFlags">optional Int32 lFlags = 0</param>
		[SupportByVersion("MSHTML", 4)]
		public virtual object getAttribute(string strAttributeName, object lFlags)
		{
			return InvokerService.InvokeInternal.ExecuteVariantMethodGet(this, "getAttribute", strAttributeName, lFlags);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="strAttributeName">string strAttributeName</param>
		[CustomMethod]
		[SupportByVersion("MSHTML", 4)]
		public virtual object getAttribute(string strAttributeName)
		{
			return InvokerService.InvokeInternal.ExecuteVariantMethodGet(this, "getAttribute", strAttributeName);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="strAttributeName">string strAttributeName</param>
		/// <param name="lFlags">optional Int32 lFlags = 1</param>
		[SupportByVersion("MSHTML", 4)]
		public virtual bool removeAttribute(string strAttributeName, object lFlags)
		{
			return InvokerService.InvokeInternal.ExecuteBoolMethodGet(this, "removeAttribute", strAttributeName, lFlags);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="strAttributeName">string strAttributeName</param>
		[CustomMethod]
		[SupportByVersion("MSHTML", 4)]
		public virtual bool removeAttribute(string strAttributeName)
		{
			return InvokerService.InvokeInternal.ExecuteBoolMethodGet(this, "removeAttribute", strAttributeName);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="varargStart">optional object varargStart</param>
		[SupportByVersion("MSHTML", 4)]
		public virtual void scrollIntoView(object varargStart)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "scrollIntoView", varargStart);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[CustomMethod]
		[SupportByVersion("MSHTML", 4)]
		public virtual void scrollIntoView()
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "scrollIntoView");
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="pChild">NetOffice.MSHTMLApi.IHTMLElement pChild</param>
		[SupportByVersion("MSHTML", 4)]
		public virtual bool contains(NetOffice.MSHTMLApi.IHTMLElement pChild)
		{
			return InvokerService.InvokeInternal.ExecuteBoolMethodGet(this, "contains", pChild);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="where">string where</param>
		/// <param name="html">string html</param>
		[SupportByVersion("MSHTML", 4)]
		public virtual void insertAdjacentHTML(string where, string html)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "insertAdjacentHTML", where, html);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="where">string where</param>
		/// <param name="text">string text</param>
		[SupportByVersion("MSHTML", 4)]
		public virtual void insertAdjacentText(string where, string text)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "insertAdjacentText", where, text);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public virtual void click()
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "click");
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
		/// <param name="containerCapture">optional bool containerCapture = true</param>
		[SupportByVersion("MSHTML", 4)]
		public virtual void setCapture(object containerCapture)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "setCapture", containerCapture);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[CustomMethod]
		[SupportByVersion("MSHTML", 4)]
		public virtual void setCapture()
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "setCapture");
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public virtual void releaseCapture()
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "releaseCapture");
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="x">Int32 x</param>
		/// <param name="y">Int32 y</param>
		[SupportByVersion("MSHTML", 4)]
		public virtual string componentFromPoint(Int32 x, Int32 y)
		{
			return InvokerService.InvokeInternal.ExecuteStringMethodGet(this, "componentFromPoint", x, y);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="component">optional object component</param>
		[SupportByVersion("MSHTML", 4)]
		public virtual void doScroll(object component)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "doScroll", component);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[CustomMethod]
		[SupportByVersion("MSHTML", 4)]
		public virtual void doScroll()
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "doScroll");
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public virtual NetOffice.MSHTMLApi.IHTMLRectCollection getClientRects()
		{
			return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.MSHTMLApi.IHTMLRectCollection>(this, "getClientRects", typeof(NetOffice.MSHTMLApi.IHTMLRectCollection));
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public virtual NetOffice.MSHTMLApi.IHTMLRect getBoundingClientRect()
		{
			return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.MSHTMLApi.IHTMLRect>(this, "getBoundingClientRect", typeof(NetOffice.MSHTMLApi.IHTMLRect));
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="propname">string propname</param>
		/// <param name="expression">string expression</param>
		/// <param name="language">optional string language = </param>
		[SupportByVersion("MSHTML", 4)]
		public virtual void setExpression(string propname, string expression, object language)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "setExpression", propname, expression, language);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="propname">string propname</param>
		/// <param name="expression">string expression</param>
		[CustomMethod]
		[SupportByVersion("MSHTML", 4)]
		public virtual void setExpression(string propname, string expression)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "setExpression", propname, expression);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="propname">string propname</param>
		[SupportByVersion("MSHTML", 4)]
		public virtual object getExpression(string propname)
		{
			return InvokerService.InvokeInternal.ExecuteVariantMethodGet(this, "getExpression", propname);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="propname">string propname</param>
		[SupportByVersion("MSHTML", 4)]
		public virtual bool removeExpression(string propname)
		{
			return InvokerService.InvokeInternal.ExecuteBoolMethodGet(this, "removeExpression", propname);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public virtual void focus()
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "focus");
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public virtual void blur()
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "blur");
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="pUnk">object pUnk</param>
		[SupportByVersion("MSHTML", 4)]
		public virtual void addFilter(object pUnk)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "addFilter", pUnk);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="pUnk">object pUnk</param>
		[SupportByVersion("MSHTML", 4)]
		public virtual void removeFilter(object pUnk)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "removeFilter", pUnk);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="_event">string event</param>
		/// <param name="pdisp">object pdisp</param>
		[SupportByVersion("MSHTML", 4)]
		public virtual bool attachEvent(string _event, object pdisp)
		{
			return InvokerService.InvokeInternal.ExecuteBoolMethodGet(this, "attachEvent", _event, pdisp);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="_event">string event</param>
		/// <param name="pdisp">object pdisp</param>
		[SupportByVersion("MSHTML", 4)]
		public virtual void detachEvent(string _event, object pdisp)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "detachEvent", _event, pdisp);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public virtual object createControlRange()
		{
			return InvokerService.InvokeInternal.ExecuteVariantMethodGet(this, "createControlRange");
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public virtual void clearAttributes()
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "clearAttributes");
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="where">string where</param>
		/// <param name="insertedElement">NetOffice.MSHTMLApi.IHTMLElement insertedElement</param>
		[SupportByVersion("MSHTML", 4)]
		[BaseResult]
		public virtual NetOffice.MSHTMLApi.IHTMLElement insertAdjacentElement(string where, NetOffice.MSHTMLApi.IHTMLElement insertedElement)
		{
			return InvokerService.InvokeInternal.ExecuteBaseReferenceMethodGet<NetOffice.MSHTMLApi.IHTMLElement>(this, "insertAdjacentElement", where, insertedElement);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="apply">NetOffice.MSHTMLApi.IHTMLElement apply</param>
		/// <param name="where">string where</param>
		[SupportByVersion("MSHTML", 4)]
		[BaseResult]
		public virtual NetOffice.MSHTMLApi.IHTMLElement applyElement(NetOffice.MSHTMLApi.IHTMLElement apply, string where)
		{
			return InvokerService.InvokeInternal.ExecuteBaseReferenceMethodGet<NetOffice.MSHTMLApi.IHTMLElement>(this, "applyElement", apply, where);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="where">string where</param>
		[SupportByVersion("MSHTML", 4)]
		public virtual string getAdjacentText(string where)
		{
			return InvokerService.InvokeInternal.ExecuteStringMethodGet(this, "getAdjacentText", where);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="where">string where</param>
		/// <param name="newText">string newText</param>
		[SupportByVersion("MSHTML", 4)]
		public virtual string replaceAdjacentText(string where, string newText)
		{
			return InvokerService.InvokeInternal.ExecuteStringMethodGet(this, "replaceAdjacentText", where, newText);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="bstrUrl">string bstrUrl</param>
		/// <param name="pvarFactory">optional object pvarFactory</param>
		[SupportByVersion("MSHTML", 4)]
		public virtual Int32 addBehavior(string bstrUrl, object pvarFactory)
		{
			return InvokerService.InvokeInternal.ExecuteInt32MethodGet(this, "addBehavior", bstrUrl, pvarFactory);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="bstrUrl">string bstrUrl</param>
		[CustomMethod]
		[SupportByVersion("MSHTML", 4)]
		public virtual Int32 addBehavior(string bstrUrl)
		{
			return InvokerService.InvokeInternal.ExecuteInt32MethodGet(this, "addBehavior", bstrUrl);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="cookie">Int32 cookie</param>
		[SupportByVersion("MSHTML", 4)]
		public virtual bool removeBehavior(Int32 cookie)
		{
			return InvokerService.InvokeInternal.ExecuteBoolMethodGet(this, "removeBehavior", cookie);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="v">string v</param>
		[SupportByVersion("MSHTML", 4)]
		[BaseResult]
		public virtual NetOffice.MSHTMLApi.IHTMLElementCollection getElementsByTagName(string v)
		{
			return InvokerService.InvokeInternal.ExecuteBaseReferenceMethodGet<NetOffice.MSHTMLApi.IHTMLElementCollection>(this, "getElementsByTagName", v);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="mergeThis">NetOffice.MSHTMLApi.IHTMLElement mergeThis</param>
		/// <param name="pvarFlags">optional object pvarFlags</param>
		[SupportByVersion("MSHTML", 4)]
		public virtual void mergeAttributes(NetOffice.MSHTMLApi.IHTMLElement mergeThis, object pvarFlags)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "mergeAttributes", mergeThis, pvarFlags);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="mergeThis">NetOffice.MSHTMLApi.IHTMLElement mergeThis</param>
		[CustomMethod]
		[SupportByVersion("MSHTML", 4)]
		public virtual void mergeAttributes(NetOffice.MSHTMLApi.IHTMLElement mergeThis)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "mergeAttributes", mergeThis);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public virtual void setActive()
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "setActive");
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="bstrEventName">string bstrEventName</param>
		/// <param name="pvarEventObject">optional object pvarEventObject</param>
		[SupportByVersion("MSHTML", 4)]
		public virtual bool FireEvent(string bstrEventName, object pvarEventObject)
		{
			return InvokerService.InvokeInternal.ExecuteBoolMethodGet(this, "FireEvent", bstrEventName, pvarEventObject);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="bstrEventName">string bstrEventName</param>
		[CustomMethod]
		[SupportByVersion("MSHTML", 4)]
		public virtual bool FireEvent(string bstrEventName)
		{
			return InvokerService.InvokeInternal.ExecuteBoolMethodGet(this, "FireEvent", bstrEventName);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public virtual bool dragDrop()
		{
			return InvokerService.InvokeInternal.ExecuteBoolMethodGet(this, "dragDrop");
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public virtual void normalize()
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "normalize");
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="bstrName">string bstrName</param>
		[SupportByVersion("MSHTML", 4)]
		[BaseResult]
		public virtual NetOffice.MSHTMLApi.IHTMLDOMAttribute getAttributeNode(string bstrName)
		{
			return InvokerService.InvokeInternal.ExecuteBaseReferenceMethodGet<NetOffice.MSHTMLApi.IHTMLDOMAttribute>(this, "getAttributeNode", bstrName);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="pattr">NetOffice.MSHTMLApi.IHTMLDOMAttribute pattr</param>
		[SupportByVersion("MSHTML", 4)]
		[BaseResult]
		public virtual NetOffice.MSHTMLApi.IHTMLDOMAttribute setAttributeNode(NetOffice.MSHTMLApi.IHTMLDOMAttribute pattr)
		{
			return InvokerService.InvokeInternal.ExecuteBaseReferenceMethodGet<NetOffice.MSHTMLApi.IHTMLDOMAttribute>(this, "setAttributeNode", pattr);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="pattr">NetOffice.MSHTMLApi.IHTMLDOMAttribute pattr</param>
		[SupportByVersion("MSHTML", 4)]
		[BaseResult]
		public virtual NetOffice.MSHTMLApi.IHTMLDOMAttribute removeAttributeNode(NetOffice.MSHTMLApi.IHTMLDOMAttribute pattr)
		{
			return InvokerService.InvokeInternal.ExecuteBaseReferenceMethodGet<NetOffice.MSHTMLApi.IHTMLDOMAttribute>(this, "removeAttributeNode", pattr);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public virtual bool hasChildNodes()
		{
			return InvokerService.InvokeInternal.ExecuteBoolMethodGet(this, "hasChildNodes");
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="newChild">NetOffice.MSHTMLApi.IHTMLDOMNode newChild</param>
		/// <param name="refChild">optional object refChild</param>
		[SupportByVersion("MSHTML", 4)]
		[BaseResult]
		public virtual NetOffice.MSHTMLApi.IHTMLDOMNode insertBefore(NetOffice.MSHTMLApi.IHTMLDOMNode newChild, object refChild)
		{
			return InvokerService.InvokeInternal.ExecuteBaseReferenceMethodGet<NetOffice.MSHTMLApi.IHTMLDOMNode>(this, "insertBefore", newChild, refChild);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="newChild">NetOffice.MSHTMLApi.IHTMLDOMNode newChild</param>
		[CustomMethod]
		[BaseResult]
		[SupportByVersion("MSHTML", 4)]
		public virtual NetOffice.MSHTMLApi.IHTMLDOMNode insertBefore(NetOffice.MSHTMLApi.IHTMLDOMNode newChild)
		{
			return InvokerService.InvokeInternal.ExecuteBaseReferenceMethodGet<NetOffice.MSHTMLApi.IHTMLDOMNode>(this, "insertBefore", newChild);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="oldChild">NetOffice.MSHTMLApi.IHTMLDOMNode oldChild</param>
		[SupportByVersion("MSHTML", 4)]
		[BaseResult]
		public virtual NetOffice.MSHTMLApi.IHTMLDOMNode removeChild(NetOffice.MSHTMLApi.IHTMLDOMNode oldChild)
		{
			return InvokerService.InvokeInternal.ExecuteBaseReferenceMethodGet<NetOffice.MSHTMLApi.IHTMLDOMNode>(this, "removeChild", oldChild);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="newChild">NetOffice.MSHTMLApi.IHTMLDOMNode newChild</param>
		/// <param name="oldChild">NetOffice.MSHTMLApi.IHTMLDOMNode oldChild</param>
		[SupportByVersion("MSHTML", 4)]
		[BaseResult]
		public virtual NetOffice.MSHTMLApi.IHTMLDOMNode replaceChild(NetOffice.MSHTMLApi.IHTMLDOMNode newChild, NetOffice.MSHTMLApi.IHTMLDOMNode oldChild)
		{
			return InvokerService.InvokeInternal.ExecuteBaseReferenceMethodGet<NetOffice.MSHTMLApi.IHTMLDOMNode>(this, "replaceChild", newChild, oldChild);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="fDeep">bool fDeep</param>
		[SupportByVersion("MSHTML", 4)]
		[BaseResult]
		public virtual NetOffice.MSHTMLApi.IHTMLDOMNode cloneNode(bool fDeep)
		{
			return InvokerService.InvokeInternal.ExecuteBaseReferenceMethodGet<NetOffice.MSHTMLApi.IHTMLDOMNode>(this, "cloneNode", fDeep);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="fDeep">optional bool fDeep = false</param>
		[SupportByVersion("MSHTML", 4)]
		[BaseResult]
		public virtual NetOffice.MSHTMLApi.IHTMLDOMNode removeNode(object fDeep)
		{
			return InvokerService.InvokeInternal.ExecuteBaseReferenceMethodGet<NetOffice.MSHTMLApi.IHTMLDOMNode>(this, "removeNode", fDeep);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[CustomMethod]
		[BaseResult]
		[SupportByVersion("MSHTML", 4)]
		public virtual NetOffice.MSHTMLApi.IHTMLDOMNode removeNode()
		{
			return InvokerService.InvokeInternal.ExecuteBaseReferenceMethodGet<NetOffice.MSHTMLApi.IHTMLDOMNode>(this, "removeNode");
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="otherNode">NetOffice.MSHTMLApi.IHTMLDOMNode otherNode</param>
		[SupportByVersion("MSHTML", 4)]
		[BaseResult]
		public virtual NetOffice.MSHTMLApi.IHTMLDOMNode swapNode(NetOffice.MSHTMLApi.IHTMLDOMNode otherNode)
		{
			return InvokerService.InvokeInternal.ExecuteBaseReferenceMethodGet<NetOffice.MSHTMLApi.IHTMLDOMNode>(this, "swapNode", otherNode);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="replacement">NetOffice.MSHTMLApi.IHTMLDOMNode replacement</param>
		[SupportByVersion("MSHTML", 4)]
		[BaseResult]
		public virtual NetOffice.MSHTMLApi.IHTMLDOMNode replaceNode(NetOffice.MSHTMLApi.IHTMLDOMNode replacement)
		{
			return InvokerService.InvokeInternal.ExecuteBaseReferenceMethodGet<NetOffice.MSHTMLApi.IHTMLDOMNode>(this, "replaceNode", replacement);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="newChild">NetOffice.MSHTMLApi.IHTMLDOMNode newChild</param>
		[SupportByVersion("MSHTML", 4)]
		[BaseResult]
		public virtual NetOffice.MSHTMLApi.IHTMLDOMNode appendChild(NetOffice.MSHTMLApi.IHTMLDOMNode newChild)
		{
			return InvokerService.InvokeInternal.ExecuteBaseReferenceMethodGet<NetOffice.MSHTMLApi.IHTMLDOMNode>(this, "appendChild", newChild);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="bstrName">string bstrName</param>
		[SupportByVersion("MSHTML", 4)]
		[BaseResult]
		public virtual NetOffice.MSHTMLApi.IHTMLDOMAttribute2 ie8_getAttributeNode(string bstrName)
		{
			return InvokerService.InvokeInternal.ExecuteBaseReferenceMethodGet<NetOffice.MSHTMLApi.IHTMLDOMAttribute2>(this, "ie8_getAttributeNode", bstrName);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="pattr">NetOffice.MSHTMLApi.IHTMLDOMAttribute2 pattr</param>
		[SupportByVersion("MSHTML", 4)]
		[BaseResult]
		public virtual NetOffice.MSHTMLApi.IHTMLDOMAttribute2 ie8_setAttributeNode(NetOffice.MSHTMLApi.IHTMLDOMAttribute2 pattr)
		{
			return InvokerService.InvokeInternal.ExecuteBaseReferenceMethodGet<NetOffice.MSHTMLApi.IHTMLDOMAttribute2>(this, "ie8_setAttributeNode", pattr);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="pattr">NetOffice.MSHTMLApi.IHTMLDOMAttribute2 pattr</param>
		[SupportByVersion("MSHTML", 4)]
		[BaseResult]
		public virtual NetOffice.MSHTMLApi.IHTMLDOMAttribute2 ie8_removeAttributeNode(NetOffice.MSHTMLApi.IHTMLDOMAttribute2 pattr)
		{
			return InvokerService.InvokeInternal.ExecuteBaseReferenceMethodGet<NetOffice.MSHTMLApi.IHTMLDOMAttribute2>(this, "ie8_removeAttributeNode", pattr);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="name">string name</param>
		[SupportByVersion("MSHTML", 4)]
		public virtual bool hasAttribute(string name)
		{
			return InvokerService.InvokeInternal.ExecuteBoolMethodGet(this, "hasAttribute", name);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="strAttributeName">string strAttributeName</param>
		[SupportByVersion("MSHTML", 4)]
		public virtual object ie8_getAttribute(string strAttributeName)
		{
			return InvokerService.InvokeInternal.ExecuteVariantMethodGet(this, "ie8_getAttribute", strAttributeName);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="strAttributeName">string strAttributeName</param>
		/// <param name="attributeValue">object attributeValue</param>
		[SupportByVersion("MSHTML", 4)]
		public virtual void ie8_setAttribute(string strAttributeName, object attributeValue)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "ie8_setAttribute", strAttributeName, attributeValue);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="strAttributeName">string strAttributeName</param>
		[SupportByVersion("MSHTML", 4)]
		public virtual bool ie8_removeAttribute(string strAttributeName)
		{
			return InvokerService.InvokeInternal.ExecuteBoolMethodGet(this, "ie8_removeAttribute", strAttributeName);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public virtual bool hasAttributes()
		{
			return InvokerService.InvokeInternal.ExecuteBoolMethodGet(this, "hasAttributes");
		}

		#endregion

		#pragma warning restore
	}
}


