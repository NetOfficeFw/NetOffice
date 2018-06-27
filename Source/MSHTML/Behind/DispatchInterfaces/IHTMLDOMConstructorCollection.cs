using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;
using NetOffice.MSHTMLApi;

namespace NetOffice.MSHTMLApi.Behind
{
	/// <summary>
	/// DispatchInterface IHTMLDOMConstructorCollection 
	/// SupportByVersion MSHTML, 4
	/// </summary>
	[SupportByVersion("MSHTML", 4)]
	[EntityType(EntityType.IsDispatchInterface)]
 	public class IHTMLDOMConstructorCollection : COMObject, NetOffice.MSHTMLApi.IHTMLDOMConstructorCollection
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
                    _contractType = typeof(NetOffice.MSHTMLApi.IHTMLDOMConstructorCollection);
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
                    _type = typeof(IHTMLDOMConstructorCollection);
                return _type;
            }
        }
        
        #endregion
        
		#region Ctor

		/// <summary>
		/// Stub Ctor, not indented to use
		/// </summary>
		public IHTMLDOMConstructorCollection() : base()
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
		public virtual object Attr
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteReferencePropertyGet(this, "Attr");
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		[SupportByVersion("MSHTML", 4), ProxyResult]
		public virtual object BehaviorUrnsCollection
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteReferencePropertyGet(this, "BehaviorUrnsCollection");
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		[SupportByVersion("MSHTML", 4), ProxyResult]
		public virtual object BookmarkCollection
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteReferencePropertyGet(this, "BookmarkCollection");
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		[SupportByVersion("MSHTML", 4), ProxyResult]
		public virtual object compatibleInfo
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteReferencePropertyGet(this, "compatibleInfo");
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		[SupportByVersion("MSHTML", 4), ProxyResult]
		public virtual object CompatibleInfoCollection
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteReferencePropertyGet(this, "CompatibleInfoCollection");
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		[SupportByVersion("MSHTML", 4), ProxyResult]
		public virtual object ControlRangeCollection
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteReferencePropertyGet(this, "ControlRangeCollection");
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		[SupportByVersion("MSHTML", 4), ProxyResult]
		public virtual object CSSCurrentStyleDeclaration
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteReferencePropertyGet(this, "CSSCurrentStyleDeclaration");
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		[SupportByVersion("MSHTML", 4), ProxyResult]
		public virtual object CSSRuleList
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteReferencePropertyGet(this, "CSSRuleList");
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		[SupportByVersion("MSHTML", 4), ProxyResult]
		public virtual object CSSRuleStyleDeclaration
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteReferencePropertyGet(this, "CSSRuleStyleDeclaration");
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		[SupportByVersion("MSHTML", 4), ProxyResult]
		public virtual object CSSStyleDeclaration
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteReferencePropertyGet(this, "CSSStyleDeclaration");
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		[SupportByVersion("MSHTML", 4), ProxyResult]
		public virtual object CSSStyleRule
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteReferencePropertyGet(this, "CSSStyleRule");
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		[SupportByVersion("MSHTML", 4), ProxyResult]
		public virtual object CSSStyleSheet
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteReferencePropertyGet(this, "CSSStyleSheet");
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		[SupportByVersion("MSHTML", 4), ProxyResult]
		public virtual object dataTransfer
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteReferencePropertyGet(this, "dataTransfer");
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		[SupportByVersion("MSHTML", 4), ProxyResult]
		public virtual object DOMImplementation
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteReferencePropertyGet(this, "DOMImplementation");
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		[SupportByVersion("MSHTML", 4), ProxyResult]
		public virtual object element
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteReferencePropertyGet(this, "element");
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		[SupportByVersion("MSHTML", 4), ProxyResult]
		public virtual object get_event()
		{
			return InvokerService.InvokeInternal.ExecuteReferencePropertyGet(this, "event");
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		[SupportByVersion("MSHTML", 4), ProxyResult]
		public virtual object history
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteReferencePropertyGet(this, "history");
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		[SupportByVersion("MSHTML", 4), ProxyResult]
		public virtual object HTCElementBehaviorDefaults
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteReferencePropertyGet(this, "HTCElementBehaviorDefaults");
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		[SupportByVersion("MSHTML", 4), ProxyResult]
		public virtual object HTMLAnchorElement
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteReferencePropertyGet(this, "HTMLAnchorElement");
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		[SupportByVersion("MSHTML", 4), ProxyResult]
		public virtual object HTMLAreaElement
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteReferencePropertyGet(this, "HTMLAreaElement");
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		[SupportByVersion("MSHTML", 4), ProxyResult]
		public virtual object HTMLAreasCollection
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteReferencePropertyGet(this, "HTMLAreasCollection");
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		[SupportByVersion("MSHTML", 4), ProxyResult]
		public virtual object HTMLBaseElement
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteReferencePropertyGet(this, "HTMLBaseElement");
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		[SupportByVersion("MSHTML", 4), ProxyResult]
		public virtual object HTMLBaseFontElement
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteReferencePropertyGet(this, "HTMLBaseFontElement");
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		[SupportByVersion("MSHTML", 4), ProxyResult]
		public virtual object HTMLBGSoundElement
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteReferencePropertyGet(this, "HTMLBGSoundElement");
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		[SupportByVersion("MSHTML", 4), ProxyResult]
		public virtual object HTMLBlockElement
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteReferencePropertyGet(this, "HTMLBlockElement");
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		[SupportByVersion("MSHTML", 4), ProxyResult]
		public virtual object HTMLBodyElement
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteReferencePropertyGet(this, "HTMLBodyElement");
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		[SupportByVersion("MSHTML", 4), ProxyResult]
		public virtual object HTMLBRElement
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteReferencePropertyGet(this, "HTMLBRElement");
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		[SupportByVersion("MSHTML", 4), ProxyResult]
		public virtual object HTMLButtonElement
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteReferencePropertyGet(this, "HTMLButtonElement");
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		[SupportByVersion("MSHTML", 4), ProxyResult]
		public virtual object HTMLCollection
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteReferencePropertyGet(this, "HTMLCollection");
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		[SupportByVersion("MSHTML", 4), ProxyResult]
		public virtual object HTMLCommentElement
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteReferencePropertyGet(this, "HTMLCommentElement");
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		[SupportByVersion("MSHTML", 4), ProxyResult]
		public virtual object HTMLDDElement
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteReferencePropertyGet(this, "HTMLDDElement");
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		[SupportByVersion("MSHTML", 4), ProxyResult]
		public virtual object HTMLDivElement
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteReferencePropertyGet(this, "HTMLDivElement");
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		[SupportByVersion("MSHTML", 4), ProxyResult]
		public virtual object HTMLDocument
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteReferencePropertyGet(this, "HTMLDocument");
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		[SupportByVersion("MSHTML", 4), ProxyResult]
		public virtual object HTMLDListElement
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteReferencePropertyGet(this, "HTMLDListElement");
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		[SupportByVersion("MSHTML", 4), ProxyResult]
		public virtual object HTMLDTElement
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteReferencePropertyGet(this, "HTMLDTElement");
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		[SupportByVersion("MSHTML", 4), ProxyResult]
		public virtual object HTMLEmbedElement
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteReferencePropertyGet(this, "HTMLEmbedElement");
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		[SupportByVersion("MSHTML", 4), ProxyResult]
		public virtual object HTMLFieldSetElement
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteReferencePropertyGet(this, "HTMLFieldSetElement");
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		[SupportByVersion("MSHTML", 4), ProxyResult]
		public virtual object HTMLFontElement
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteReferencePropertyGet(this, "HTMLFontElement");
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		[SupportByVersion("MSHTML", 4), ProxyResult]
		public virtual object HTMLFormElement
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteReferencePropertyGet(this, "HTMLFormElement");
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		[SupportByVersion("MSHTML", 4), ProxyResult]
		public virtual object HTMLFrameElement
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteReferencePropertyGet(this, "HTMLFrameElement");
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		[SupportByVersion("MSHTML", 4), ProxyResult]
		public virtual object HTMLFrameSetElement
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteReferencePropertyGet(this, "HTMLFrameSetElement");
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		[SupportByVersion("MSHTML", 4), ProxyResult]
		public virtual object HTMLGenericElement
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteReferencePropertyGet(this, "HTMLGenericElement");
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		[SupportByVersion("MSHTML", 4), ProxyResult]
		public virtual object HTMLHeadElement
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteReferencePropertyGet(this, "HTMLHeadElement");
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		[SupportByVersion("MSHTML", 4), ProxyResult]
		public virtual object HTMLHeadingElement
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteReferencePropertyGet(this, "HTMLHeadingElement");
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		[SupportByVersion("MSHTML", 4), ProxyResult]
		public virtual object HTMLHRElement
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteReferencePropertyGet(this, "HTMLHRElement");
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		[SupportByVersion("MSHTML", 4), ProxyResult]
		public virtual object HTMLHtmlElement
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteReferencePropertyGet(this, "HTMLHtmlElement");
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		[SupportByVersion("MSHTML", 4), ProxyResult]
		public virtual object HTMLIFrameElement
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteReferencePropertyGet(this, "HTMLIFrameElement");
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		[SupportByVersion("MSHTML", 4), ProxyResult]
		public virtual object HTMLImageElement
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteReferencePropertyGet(this, "HTMLImageElement");
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		[SupportByVersion("MSHTML", 4), ProxyResult]
		public virtual object HTMLInputElement
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteReferencePropertyGet(this, "HTMLInputElement");
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		[SupportByVersion("MSHTML", 4), ProxyResult]
		public virtual object HTMLIsIndexElement
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteReferencePropertyGet(this, "HTMLIsIndexElement");
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		[SupportByVersion("MSHTML", 4), ProxyResult]
		public virtual object HTMLLabelElement
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteReferencePropertyGet(this, "HTMLLabelElement");
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		[SupportByVersion("MSHTML", 4), ProxyResult]
		public virtual object HTMLLegendElement
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteReferencePropertyGet(this, "HTMLLegendElement");
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		[SupportByVersion("MSHTML", 4), ProxyResult]
		public virtual object HTMLLIElement
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteReferencePropertyGet(this, "HTMLLIElement");
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		[SupportByVersion("MSHTML", 4), ProxyResult]
		public virtual object HTMLLinkElement
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteReferencePropertyGet(this, "HTMLLinkElement");
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		[SupportByVersion("MSHTML", 4), ProxyResult]
		public virtual object HTMLMapElement
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteReferencePropertyGet(this, "HTMLMapElement");
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		[SupportByVersion("MSHTML", 4), ProxyResult]
		public virtual object HTMLMarqueeElement
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteReferencePropertyGet(this, "HTMLMarqueeElement");
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		[SupportByVersion("MSHTML", 4), ProxyResult]
		public virtual object HTMLMetaElement
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteReferencePropertyGet(this, "HTMLMetaElement");
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		[SupportByVersion("MSHTML", 4), ProxyResult]
		public virtual object HTMLModelessDialog
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteReferencePropertyGet(this, "HTMLModelessDialog");
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		[SupportByVersion("MSHTML", 4), ProxyResult]
		public virtual object HTMLNamespaceInfo
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteReferencePropertyGet(this, "HTMLNamespaceInfo");
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		[SupportByVersion("MSHTML", 4), ProxyResult]
		public virtual object HTMLNamespaceInfoCollection
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteReferencePropertyGet(this, "HTMLNamespaceInfoCollection");
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		[SupportByVersion("MSHTML", 4), ProxyResult]
		public virtual object HTMLNextIdElement
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteReferencePropertyGet(this, "HTMLNextIdElement");
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		[SupportByVersion("MSHTML", 4), ProxyResult]
		public virtual object HTMLNoShowElement
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteReferencePropertyGet(this, "HTMLNoShowElement");
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		[SupportByVersion("MSHTML", 4), ProxyResult]
		public virtual object HTMLObjectElement
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteReferencePropertyGet(this, "HTMLObjectElement");
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		[SupportByVersion("MSHTML", 4), ProxyResult]
		public virtual object HTMLOListElement
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteReferencePropertyGet(this, "HTMLOListElement");
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		[SupportByVersion("MSHTML", 4), ProxyResult]
		public virtual object HTMLOptionElement
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteReferencePropertyGet(this, "HTMLOptionElement");
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		[SupportByVersion("MSHTML", 4), ProxyResult]
		public virtual object HTMLParagraphElement
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteReferencePropertyGet(this, "HTMLParagraphElement");
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		[SupportByVersion("MSHTML", 4), ProxyResult]
		public virtual object HTMLParamElement
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteReferencePropertyGet(this, "HTMLParamElement");
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		[SupportByVersion("MSHTML", 4), ProxyResult]
		public virtual object HTMLPhraseElement
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteReferencePropertyGet(this, "HTMLPhraseElement");
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		[SupportByVersion("MSHTML", 4), ProxyResult]
		public virtual object HTMLPluginsCollection
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteReferencePropertyGet(this, "HTMLPluginsCollection");
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		[SupportByVersion("MSHTML", 4), ProxyResult]
		public virtual object HTMLPopup
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteReferencePropertyGet(this, "HTMLPopup");
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		[SupportByVersion("MSHTML", 4), ProxyResult]
		public virtual object HTMLScriptElement
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteReferencePropertyGet(this, "HTMLScriptElement");
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		[SupportByVersion("MSHTML", 4), ProxyResult]
		public virtual object HTMLSelectElement
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteReferencePropertyGet(this, "HTMLSelectElement");
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		[SupportByVersion("MSHTML", 4), ProxyResult]
		public virtual object HTMLSpanElement
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteReferencePropertyGet(this, "HTMLSpanElement");
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		[SupportByVersion("MSHTML", 4), ProxyResult]
		public virtual object HTMLStyleElement
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteReferencePropertyGet(this, "HTMLStyleElement");
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		[SupportByVersion("MSHTML", 4), ProxyResult]
		public virtual object HTMLTableCaptionElement
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteReferencePropertyGet(this, "HTMLTableCaptionElement");
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		[SupportByVersion("MSHTML", 4), ProxyResult]
		public virtual object HTMLTableCellElement
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteReferencePropertyGet(this, "HTMLTableCellElement");
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		[SupportByVersion("MSHTML", 4), ProxyResult]
		public virtual object HTMLTableColElement
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteReferencePropertyGet(this, "HTMLTableColElement");
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		[SupportByVersion("MSHTML", 4), ProxyResult]
		public virtual object HTMLTableElement
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteReferencePropertyGet(this, "HTMLTableElement");
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		[SupportByVersion("MSHTML", 4), ProxyResult]
		public virtual object HTMLTableRowElement
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteReferencePropertyGet(this, "HTMLTableRowElement");
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		[SupportByVersion("MSHTML", 4), ProxyResult]
		public virtual object HTMLTableSectionElement
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteReferencePropertyGet(this, "HTMLTableSectionElement");
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		[SupportByVersion("MSHTML", 4), ProxyResult]
		public virtual object HTMLTextAreaElement
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteReferencePropertyGet(this, "HTMLTextAreaElement");
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		[SupportByVersion("MSHTML", 4), ProxyResult]
		public virtual object HTMLTextElement
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteReferencePropertyGet(this, "HTMLTextElement");
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		[SupportByVersion("MSHTML", 4), ProxyResult]
		public virtual object HTMLTitleElement
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteReferencePropertyGet(this, "HTMLTitleElement");
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		[SupportByVersion("MSHTML", 4), ProxyResult]
		public virtual object HTMLUListElement
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteReferencePropertyGet(this, "HTMLUListElement");
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		[SupportByVersion("MSHTML", 4), ProxyResult]
		public virtual object HTMLUnknownElement
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteReferencePropertyGet(this, "HTMLUnknownElement");
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		[SupportByVersion("MSHTML", 4), ProxyResult]
		public virtual object Image
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteReferencePropertyGet(this, "Image");
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		[SupportByVersion("MSHTML", 4), ProxyResult]
		public virtual object location
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteReferencePropertyGet(this, "location");
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		[SupportByVersion("MSHTML", 4), ProxyResult]
		public virtual object NamedNodeMap
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteReferencePropertyGet(this, "NamedNodeMap");
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		[SupportByVersion("MSHTML", 4), ProxyResult]
		public virtual object navigator
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteReferencePropertyGet(this, "navigator");
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		[SupportByVersion("MSHTML", 4), ProxyResult]
		public virtual object NodeList
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteReferencePropertyGet(this, "NodeList");
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		[SupportByVersion("MSHTML", 4), ProxyResult]
		public virtual object Option
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteReferencePropertyGet(this, "Option");
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		[SupportByVersion("MSHTML", 4), ProxyResult]
		public virtual object screen
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteReferencePropertyGet(this, "screen");
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		[SupportByVersion("MSHTML", 4), ProxyResult]
		public virtual object selection
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteReferencePropertyGet(this, "selection");
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		[SupportByVersion("MSHTML", 4), ProxyResult]
		public virtual object StaticNodeList
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteReferencePropertyGet(this, "StaticNodeList");
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		[SupportByVersion("MSHTML", 4), ProxyResult]
		public virtual object Storage
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteReferencePropertyGet(this, "Storage");
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		[SupportByVersion("MSHTML", 4), ProxyResult]
		public virtual object StyleSheetList
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteReferencePropertyGet(this, "StyleSheetList");
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		[SupportByVersion("MSHTML", 4), ProxyResult]
		public virtual object StyleSheetPage
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteReferencePropertyGet(this, "StyleSheetPage");
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		[SupportByVersion("MSHTML", 4), ProxyResult]
		public virtual object StyleSheetPageList
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteReferencePropertyGet(this, "StyleSheetPageList");
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		[SupportByVersion("MSHTML", 4), ProxyResult]
		public virtual object text
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteReferencePropertyGet(this, "text");
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		[SupportByVersion("MSHTML", 4), ProxyResult]
		public virtual object TextRange
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteReferencePropertyGet(this, "TextRange");
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		[SupportByVersion("MSHTML", 4), ProxyResult]
		public virtual object TextRangeCollection
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteReferencePropertyGet(this, "TextRangeCollection");
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		[SupportByVersion("MSHTML", 4), ProxyResult]
		public virtual object TextRectangle
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteReferencePropertyGet(this, "TextRectangle");
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		[SupportByVersion("MSHTML", 4), ProxyResult]
		public virtual object TextRectangleList
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteReferencePropertyGet(this, "TextRectangleList");
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		[SupportByVersion("MSHTML", 4), ProxyResult]
		public virtual object window
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteReferencePropertyGet(this, "window");
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		[SupportByVersion("MSHTML", 4), ProxyResult]
		public virtual object XDomainRequest
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteReferencePropertyGet(this, "XDomainRequest");
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		[SupportByVersion("MSHTML", 4), ProxyResult]
		public virtual object XMLHttpRequest
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteReferencePropertyGet(this, "XMLHttpRequest");
			}
		}

		#endregion

		#region Methods

		#endregion

		#pragma warning restore
	}
}

