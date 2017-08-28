using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.MSHTMLApi
{
	/// <summary>
	/// DispatchInterface IHTMLDOMConstructorCollection 
	/// SupportByVersion MSHTML, 4
	/// </summary>
	[SupportByVersion("MSHTML", 4)]
	[EntityType(EntityType.IsDispatchInterface)]
 	public class IHTMLDOMConstructorCollection : COMObject
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
                    _type = typeof(IHTMLDOMConstructorCollection);
                return _type;
            }
        }
        
        #endregion
        
		#region Ctor

		/// <param name="factory">current used factory core</param>
		/// <param name="parentObject">object there has created the proxy</param>
		/// <param name="proxyShare">proxy share instead if com proxy</param>
		public IHTMLDOMConstructorCollection(Core factory, ICOMObject parentObject, COMProxyShare proxyShare) : base(factory, parentObject, proxyShare)
		{
		}

		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
		public IHTMLDOMConstructorCollection(Core factory, ICOMObject parentObject, object comProxy) : base(factory, parentObject, comProxy)
		{
			
		}

        ///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public IHTMLDOMConstructorCollection(ICOMObject parentObject, object comProxy) : base(parentObject, comProxy)
		{
		}
		
		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public IHTMLDOMConstructorCollection(Core factory, ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(factory, parentObject, comProxy, comProxyType)
		{

		}

		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public IHTMLDOMConstructorCollection(ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(parentObject, comProxy, comProxyType)
		{
		}
		
		///<param name="replacedObject">object to replaced. replacedObject are not usable after this action</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public IHTMLDOMConstructorCollection(ICOMObject replacedObject) : base(replacedObject)
		{
		}
		
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public IHTMLDOMConstructorCollection() : base()
		{
		}
		
		/// <param name="progId">registered progID</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public IHTMLDOMConstructorCollection(string progId) : base(progId)
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
		public object Attr
		{
			get
			{
				return Factory.ExecuteReferencePropertyGet(this, "Attr");
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		[SupportByVersion("MSHTML", 4), ProxyResult]
		public object BehaviorUrnsCollection
		{
			get
			{
				return Factory.ExecuteReferencePropertyGet(this, "BehaviorUrnsCollection");
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		[SupportByVersion("MSHTML", 4), ProxyResult]
		public object BookmarkCollection
		{
			get
			{
				return Factory.ExecuteReferencePropertyGet(this, "BookmarkCollection");
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		[SupportByVersion("MSHTML", 4), ProxyResult]
		public object compatibleInfo
		{
			get
			{
				return Factory.ExecuteReferencePropertyGet(this, "compatibleInfo");
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		[SupportByVersion("MSHTML", 4), ProxyResult]
		public object CompatibleInfoCollection
		{
			get
			{
				return Factory.ExecuteReferencePropertyGet(this, "CompatibleInfoCollection");
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		[SupportByVersion("MSHTML", 4), ProxyResult]
		public object ControlRangeCollection
		{
			get
			{
				return Factory.ExecuteReferencePropertyGet(this, "ControlRangeCollection");
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		[SupportByVersion("MSHTML", 4), ProxyResult]
		public object CSSCurrentStyleDeclaration
		{
			get
			{
				return Factory.ExecuteReferencePropertyGet(this, "CSSCurrentStyleDeclaration");
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		[SupportByVersion("MSHTML", 4), ProxyResult]
		public object CSSRuleList
		{
			get
			{
				return Factory.ExecuteReferencePropertyGet(this, "CSSRuleList");
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		[SupportByVersion("MSHTML", 4), ProxyResult]
		public object CSSRuleStyleDeclaration
		{
			get
			{
				return Factory.ExecuteReferencePropertyGet(this, "CSSRuleStyleDeclaration");
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		[SupportByVersion("MSHTML", 4), ProxyResult]
		public object CSSStyleDeclaration
		{
			get
			{
				return Factory.ExecuteReferencePropertyGet(this, "CSSStyleDeclaration");
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		[SupportByVersion("MSHTML", 4), ProxyResult]
		public object CSSStyleRule
		{
			get
			{
				return Factory.ExecuteReferencePropertyGet(this, "CSSStyleRule");
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		[SupportByVersion("MSHTML", 4), ProxyResult]
		public object CSSStyleSheet
		{
			get
			{
				return Factory.ExecuteReferencePropertyGet(this, "CSSStyleSheet");
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		[SupportByVersion("MSHTML", 4), ProxyResult]
		public object dataTransfer
		{
			get
			{
				return Factory.ExecuteReferencePropertyGet(this, "dataTransfer");
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		[SupportByVersion("MSHTML", 4), ProxyResult]
		public object DOMImplementation
		{
			get
			{
				return Factory.ExecuteReferencePropertyGet(this, "DOMImplementation");
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		[SupportByVersion("MSHTML", 4), ProxyResult]
		public object element
		{
			get
			{
				return Factory.ExecuteReferencePropertyGet(this, "element");
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		[SupportByVersion("MSHTML", 4), ProxyResult]
		public object get_event()
		{
			return Factory.ExecuteReferencePropertyGet(this, "event");
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		[SupportByVersion("MSHTML", 4), ProxyResult]
		public object history
		{
			get
			{
				return Factory.ExecuteReferencePropertyGet(this, "history");
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		[SupportByVersion("MSHTML", 4), ProxyResult]
		public object HTCElementBehaviorDefaults
		{
			get
			{
				return Factory.ExecuteReferencePropertyGet(this, "HTCElementBehaviorDefaults");
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		[SupportByVersion("MSHTML", 4), ProxyResult]
		public object HTMLAnchorElement
		{
			get
			{
				return Factory.ExecuteReferencePropertyGet(this, "HTMLAnchorElement");
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		[SupportByVersion("MSHTML", 4), ProxyResult]
		public object HTMLAreaElement
		{
			get
			{
				return Factory.ExecuteReferencePropertyGet(this, "HTMLAreaElement");
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		[SupportByVersion("MSHTML", 4), ProxyResult]
		public object HTMLAreasCollection
		{
			get
			{
				return Factory.ExecuteReferencePropertyGet(this, "HTMLAreasCollection");
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		[SupportByVersion("MSHTML", 4), ProxyResult]
		public object HTMLBaseElement
		{
			get
			{
				return Factory.ExecuteReferencePropertyGet(this, "HTMLBaseElement");
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		[SupportByVersion("MSHTML", 4), ProxyResult]
		public object HTMLBaseFontElement
		{
			get
			{
				return Factory.ExecuteReferencePropertyGet(this, "HTMLBaseFontElement");
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		[SupportByVersion("MSHTML", 4), ProxyResult]
		public object HTMLBGSoundElement
		{
			get
			{
				return Factory.ExecuteReferencePropertyGet(this, "HTMLBGSoundElement");
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		[SupportByVersion("MSHTML", 4), ProxyResult]
		public object HTMLBlockElement
		{
			get
			{
				return Factory.ExecuteReferencePropertyGet(this, "HTMLBlockElement");
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		[SupportByVersion("MSHTML", 4), ProxyResult]
		public object HTMLBodyElement
		{
			get
			{
				return Factory.ExecuteReferencePropertyGet(this, "HTMLBodyElement");
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		[SupportByVersion("MSHTML", 4), ProxyResult]
		public object HTMLBRElement
		{
			get
			{
				return Factory.ExecuteReferencePropertyGet(this, "HTMLBRElement");
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		[SupportByVersion("MSHTML", 4), ProxyResult]
		public object HTMLButtonElement
		{
			get
			{
				return Factory.ExecuteReferencePropertyGet(this, "HTMLButtonElement");
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		[SupportByVersion("MSHTML", 4), ProxyResult]
		public object HTMLCollection
		{
			get
			{
				return Factory.ExecuteReferencePropertyGet(this, "HTMLCollection");
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		[SupportByVersion("MSHTML", 4), ProxyResult]
		public object HTMLCommentElement
		{
			get
			{
				return Factory.ExecuteReferencePropertyGet(this, "HTMLCommentElement");
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		[SupportByVersion("MSHTML", 4), ProxyResult]
		public object HTMLDDElement
		{
			get
			{
				return Factory.ExecuteReferencePropertyGet(this, "HTMLDDElement");
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		[SupportByVersion("MSHTML", 4), ProxyResult]
		public object HTMLDivElement
		{
			get
			{
				return Factory.ExecuteReferencePropertyGet(this, "HTMLDivElement");
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		[SupportByVersion("MSHTML", 4), ProxyResult]
		public object HTMLDocument
		{
			get
			{
				return Factory.ExecuteReferencePropertyGet(this, "HTMLDocument");
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		[SupportByVersion("MSHTML", 4), ProxyResult]
		public object HTMLDListElement
		{
			get
			{
				return Factory.ExecuteReferencePropertyGet(this, "HTMLDListElement");
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		[SupportByVersion("MSHTML", 4), ProxyResult]
		public object HTMLDTElement
		{
			get
			{
				return Factory.ExecuteReferencePropertyGet(this, "HTMLDTElement");
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		[SupportByVersion("MSHTML", 4), ProxyResult]
		public object HTMLEmbedElement
		{
			get
			{
				return Factory.ExecuteReferencePropertyGet(this, "HTMLEmbedElement");
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		[SupportByVersion("MSHTML", 4), ProxyResult]
		public object HTMLFieldSetElement
		{
			get
			{
				return Factory.ExecuteReferencePropertyGet(this, "HTMLFieldSetElement");
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		[SupportByVersion("MSHTML", 4), ProxyResult]
		public object HTMLFontElement
		{
			get
			{
				return Factory.ExecuteReferencePropertyGet(this, "HTMLFontElement");
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		[SupportByVersion("MSHTML", 4), ProxyResult]
		public object HTMLFormElement
		{
			get
			{
				return Factory.ExecuteReferencePropertyGet(this, "HTMLFormElement");
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		[SupportByVersion("MSHTML", 4), ProxyResult]
		public object HTMLFrameElement
		{
			get
			{
				return Factory.ExecuteReferencePropertyGet(this, "HTMLFrameElement");
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		[SupportByVersion("MSHTML", 4), ProxyResult]
		public object HTMLFrameSetElement
		{
			get
			{
				return Factory.ExecuteReferencePropertyGet(this, "HTMLFrameSetElement");
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		[SupportByVersion("MSHTML", 4), ProxyResult]
		public object HTMLGenericElement
		{
			get
			{
				return Factory.ExecuteReferencePropertyGet(this, "HTMLGenericElement");
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		[SupportByVersion("MSHTML", 4), ProxyResult]
		public object HTMLHeadElement
		{
			get
			{
				return Factory.ExecuteReferencePropertyGet(this, "HTMLHeadElement");
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		[SupportByVersion("MSHTML", 4), ProxyResult]
		public object HTMLHeadingElement
		{
			get
			{
				return Factory.ExecuteReferencePropertyGet(this, "HTMLHeadingElement");
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		[SupportByVersion("MSHTML", 4), ProxyResult]
		public object HTMLHRElement
		{
			get
			{
				return Factory.ExecuteReferencePropertyGet(this, "HTMLHRElement");
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		[SupportByVersion("MSHTML", 4), ProxyResult]
		public object HTMLHtmlElement
		{
			get
			{
				return Factory.ExecuteReferencePropertyGet(this, "HTMLHtmlElement");
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		[SupportByVersion("MSHTML", 4), ProxyResult]
		public object HTMLIFrameElement
		{
			get
			{
				return Factory.ExecuteReferencePropertyGet(this, "HTMLIFrameElement");
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		[SupportByVersion("MSHTML", 4), ProxyResult]
		public object HTMLImageElement
		{
			get
			{
				return Factory.ExecuteReferencePropertyGet(this, "HTMLImageElement");
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		[SupportByVersion("MSHTML", 4), ProxyResult]
		public object HTMLInputElement
		{
			get
			{
				return Factory.ExecuteReferencePropertyGet(this, "HTMLInputElement");
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		[SupportByVersion("MSHTML", 4), ProxyResult]
		public object HTMLIsIndexElement
		{
			get
			{
				return Factory.ExecuteReferencePropertyGet(this, "HTMLIsIndexElement");
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		[SupportByVersion("MSHTML", 4), ProxyResult]
		public object HTMLLabelElement
		{
			get
			{
				return Factory.ExecuteReferencePropertyGet(this, "HTMLLabelElement");
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		[SupportByVersion("MSHTML", 4), ProxyResult]
		public object HTMLLegendElement
		{
			get
			{
				return Factory.ExecuteReferencePropertyGet(this, "HTMLLegendElement");
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		[SupportByVersion("MSHTML", 4), ProxyResult]
		public object HTMLLIElement
		{
			get
			{
				return Factory.ExecuteReferencePropertyGet(this, "HTMLLIElement");
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		[SupportByVersion("MSHTML", 4), ProxyResult]
		public object HTMLLinkElement
		{
			get
			{
				return Factory.ExecuteReferencePropertyGet(this, "HTMLLinkElement");
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		[SupportByVersion("MSHTML", 4), ProxyResult]
		public object HTMLMapElement
		{
			get
			{
				return Factory.ExecuteReferencePropertyGet(this, "HTMLMapElement");
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		[SupportByVersion("MSHTML", 4), ProxyResult]
		public object HTMLMarqueeElement
		{
			get
			{
				return Factory.ExecuteReferencePropertyGet(this, "HTMLMarqueeElement");
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		[SupportByVersion("MSHTML", 4), ProxyResult]
		public object HTMLMetaElement
		{
			get
			{
				return Factory.ExecuteReferencePropertyGet(this, "HTMLMetaElement");
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		[SupportByVersion("MSHTML", 4), ProxyResult]
		public object HTMLModelessDialog
		{
			get
			{
				return Factory.ExecuteReferencePropertyGet(this, "HTMLModelessDialog");
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		[SupportByVersion("MSHTML", 4), ProxyResult]
		public object HTMLNamespaceInfo
		{
			get
			{
				return Factory.ExecuteReferencePropertyGet(this, "HTMLNamespaceInfo");
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		[SupportByVersion("MSHTML", 4), ProxyResult]
		public object HTMLNamespaceInfoCollection
		{
			get
			{
				return Factory.ExecuteReferencePropertyGet(this, "HTMLNamespaceInfoCollection");
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		[SupportByVersion("MSHTML", 4), ProxyResult]
		public object HTMLNextIdElement
		{
			get
			{
				return Factory.ExecuteReferencePropertyGet(this, "HTMLNextIdElement");
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		[SupportByVersion("MSHTML", 4), ProxyResult]
		public object HTMLNoShowElement
		{
			get
			{
				return Factory.ExecuteReferencePropertyGet(this, "HTMLNoShowElement");
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		[SupportByVersion("MSHTML", 4), ProxyResult]
		public object HTMLObjectElement
		{
			get
			{
				return Factory.ExecuteReferencePropertyGet(this, "HTMLObjectElement");
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		[SupportByVersion("MSHTML", 4), ProxyResult]
		public object HTMLOListElement
		{
			get
			{
				return Factory.ExecuteReferencePropertyGet(this, "HTMLOListElement");
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		[SupportByVersion("MSHTML", 4), ProxyResult]
		public object HTMLOptionElement
		{
			get
			{
				return Factory.ExecuteReferencePropertyGet(this, "HTMLOptionElement");
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		[SupportByVersion("MSHTML", 4), ProxyResult]
		public object HTMLParagraphElement
		{
			get
			{
				return Factory.ExecuteReferencePropertyGet(this, "HTMLParagraphElement");
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		[SupportByVersion("MSHTML", 4), ProxyResult]
		public object HTMLParamElement
		{
			get
			{
				return Factory.ExecuteReferencePropertyGet(this, "HTMLParamElement");
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		[SupportByVersion("MSHTML", 4), ProxyResult]
		public object HTMLPhraseElement
		{
			get
			{
				return Factory.ExecuteReferencePropertyGet(this, "HTMLPhraseElement");
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		[SupportByVersion("MSHTML", 4), ProxyResult]
		public object HTMLPluginsCollection
		{
			get
			{
				return Factory.ExecuteReferencePropertyGet(this, "HTMLPluginsCollection");
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		[SupportByVersion("MSHTML", 4), ProxyResult]
		public object HTMLPopup
		{
			get
			{
				return Factory.ExecuteReferencePropertyGet(this, "HTMLPopup");
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		[SupportByVersion("MSHTML", 4), ProxyResult]
		public object HTMLScriptElement
		{
			get
			{
				return Factory.ExecuteReferencePropertyGet(this, "HTMLScriptElement");
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		[SupportByVersion("MSHTML", 4), ProxyResult]
		public object HTMLSelectElement
		{
			get
			{
				return Factory.ExecuteReferencePropertyGet(this, "HTMLSelectElement");
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		[SupportByVersion("MSHTML", 4), ProxyResult]
		public object HTMLSpanElement
		{
			get
			{
				return Factory.ExecuteReferencePropertyGet(this, "HTMLSpanElement");
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		[SupportByVersion("MSHTML", 4), ProxyResult]
		public object HTMLStyleElement
		{
			get
			{
				return Factory.ExecuteReferencePropertyGet(this, "HTMLStyleElement");
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		[SupportByVersion("MSHTML", 4), ProxyResult]
		public object HTMLTableCaptionElement
		{
			get
			{
				return Factory.ExecuteReferencePropertyGet(this, "HTMLTableCaptionElement");
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		[SupportByVersion("MSHTML", 4), ProxyResult]
		public object HTMLTableCellElement
		{
			get
			{
				return Factory.ExecuteReferencePropertyGet(this, "HTMLTableCellElement");
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		[SupportByVersion("MSHTML", 4), ProxyResult]
		public object HTMLTableColElement
		{
			get
			{
				return Factory.ExecuteReferencePropertyGet(this, "HTMLTableColElement");
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		[SupportByVersion("MSHTML", 4), ProxyResult]
		public object HTMLTableElement
		{
			get
			{
				return Factory.ExecuteReferencePropertyGet(this, "HTMLTableElement");
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		[SupportByVersion("MSHTML", 4), ProxyResult]
		public object HTMLTableRowElement
		{
			get
			{
				return Factory.ExecuteReferencePropertyGet(this, "HTMLTableRowElement");
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		[SupportByVersion("MSHTML", 4), ProxyResult]
		public object HTMLTableSectionElement
		{
			get
			{
				return Factory.ExecuteReferencePropertyGet(this, "HTMLTableSectionElement");
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		[SupportByVersion("MSHTML", 4), ProxyResult]
		public object HTMLTextAreaElement
		{
			get
			{
				return Factory.ExecuteReferencePropertyGet(this, "HTMLTextAreaElement");
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		[SupportByVersion("MSHTML", 4), ProxyResult]
		public object HTMLTextElement
		{
			get
			{
				return Factory.ExecuteReferencePropertyGet(this, "HTMLTextElement");
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		[SupportByVersion("MSHTML", 4), ProxyResult]
		public object HTMLTitleElement
		{
			get
			{
				return Factory.ExecuteReferencePropertyGet(this, "HTMLTitleElement");
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		[SupportByVersion("MSHTML", 4), ProxyResult]
		public object HTMLUListElement
		{
			get
			{
				return Factory.ExecuteReferencePropertyGet(this, "HTMLUListElement");
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		[SupportByVersion("MSHTML", 4), ProxyResult]
		public object HTMLUnknownElement
		{
			get
			{
				return Factory.ExecuteReferencePropertyGet(this, "HTMLUnknownElement");
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		[SupportByVersion("MSHTML", 4), ProxyResult]
		public object Image
		{
			get
			{
				return Factory.ExecuteReferencePropertyGet(this, "Image");
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		[SupportByVersion("MSHTML", 4), ProxyResult]
		public object location
		{
			get
			{
				return Factory.ExecuteReferencePropertyGet(this, "location");
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		[SupportByVersion("MSHTML", 4), ProxyResult]
		public object NamedNodeMap
		{
			get
			{
				return Factory.ExecuteReferencePropertyGet(this, "NamedNodeMap");
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		[SupportByVersion("MSHTML", 4), ProxyResult]
		public object navigator
		{
			get
			{
				return Factory.ExecuteReferencePropertyGet(this, "navigator");
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		[SupportByVersion("MSHTML", 4), ProxyResult]
		public object NodeList
		{
			get
			{
				return Factory.ExecuteReferencePropertyGet(this, "NodeList");
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		[SupportByVersion("MSHTML", 4), ProxyResult]
		public object Option
		{
			get
			{
				return Factory.ExecuteReferencePropertyGet(this, "Option");
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		[SupportByVersion("MSHTML", 4), ProxyResult]
		public object screen
		{
			get
			{
				return Factory.ExecuteReferencePropertyGet(this, "screen");
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		[SupportByVersion("MSHTML", 4), ProxyResult]
		public object selection
		{
			get
			{
				return Factory.ExecuteReferencePropertyGet(this, "selection");
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		[SupportByVersion("MSHTML", 4), ProxyResult]
		public object StaticNodeList
		{
			get
			{
				return Factory.ExecuteReferencePropertyGet(this, "StaticNodeList");
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		[SupportByVersion("MSHTML", 4), ProxyResult]
		public object Storage
		{
			get
			{
				return Factory.ExecuteReferencePropertyGet(this, "Storage");
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		[SupportByVersion("MSHTML", 4), ProxyResult]
		public object StyleSheetList
		{
			get
			{
				return Factory.ExecuteReferencePropertyGet(this, "StyleSheetList");
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		[SupportByVersion("MSHTML", 4), ProxyResult]
		public object StyleSheetPage
		{
			get
			{
				return Factory.ExecuteReferencePropertyGet(this, "StyleSheetPage");
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		[SupportByVersion("MSHTML", 4), ProxyResult]
		public object StyleSheetPageList
		{
			get
			{
				return Factory.ExecuteReferencePropertyGet(this, "StyleSheetPageList");
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		[SupportByVersion("MSHTML", 4), ProxyResult]
		public object text
		{
			get
			{
				return Factory.ExecuteReferencePropertyGet(this, "text");
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		[SupportByVersion("MSHTML", 4), ProxyResult]
		public object TextRange
		{
			get
			{
				return Factory.ExecuteReferencePropertyGet(this, "TextRange");
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		[SupportByVersion("MSHTML", 4), ProxyResult]
		public object TextRangeCollection
		{
			get
			{
				return Factory.ExecuteReferencePropertyGet(this, "TextRangeCollection");
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		[SupportByVersion("MSHTML", 4), ProxyResult]
		public object TextRectangle
		{
			get
			{
				return Factory.ExecuteReferencePropertyGet(this, "TextRectangle");
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		[SupportByVersion("MSHTML", 4), ProxyResult]
		public object TextRectangleList
		{
			get
			{
				return Factory.ExecuteReferencePropertyGet(this, "TextRectangleList");
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		[SupportByVersion("MSHTML", 4), ProxyResult]
		public object window
		{
			get
			{
				return Factory.ExecuteReferencePropertyGet(this, "window");
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		[SupportByVersion("MSHTML", 4), ProxyResult]
		public object XDomainRequest
		{
			get
			{
				return Factory.ExecuteReferencePropertyGet(this, "XDomainRequest");
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		[SupportByVersion("MSHTML", 4), ProxyResult]
		public object XMLHttpRequest
		{
			get
			{
				return Factory.ExecuteReferencePropertyGet(this, "XMLHttpRequest");
			}
		}

		#endregion

		#region Methods

		#endregion

		#pragma warning restore
	}
}
