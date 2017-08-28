using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.MSHTMLApi
{
	/// <summary>
	/// DispatchInterface IHTMLElement2 
	/// SupportByVersion MSHTML, 4
	/// </summary>
	[SupportByVersion("MSHTML", 4)]
	[EntityType(EntityType.IsDispatchInterface)]
 	public class IHTMLElement2 : IHTMLElement
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
                    _type = typeof(IHTMLElement2);
                return _type;
            }
        }
        
        #endregion
        
		#region Construction

		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
		public IHTMLElement2(Core factory, ICOMObject parentObject, object comProxy) : base(factory, parentObject, comProxy)
		{
			
		}

        ///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public IHTMLElement2(ICOMObject parentObject, object comProxy) : base(parentObject, comProxy)
		{
		}
		
		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public IHTMLElement2(Core factory, ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(factory, parentObject, comProxy, comProxyType)
		{

		}

		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public IHTMLElement2(ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(parentObject, comProxy, comProxyType)
		{
		}
		
		///<param name="replacedObject">object to replaced. replacedObject are not usable after this action</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public IHTMLElement2(ICOMObject replacedObject) : base(replacedObject)
		{
		}
		
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public IHTMLElement2() : base()
		{
		}
		
		/// <param name="progId">registered progID</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public IHTMLElement2(string progId) : base(progId)
		{
		}
		
		#endregion
		
		#region Properties

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
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public NetOffice.MSHTMLApi.IHTMLCurrentStyle currentStyle
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.MSHTMLApi.IHTMLCurrentStyle>(this, "currentStyle", NetOffice.MSHTMLApi.IHTMLCurrentStyle.LateBindingApiWrapperType);
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
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public NetOffice.MSHTMLApi.IHTMLStyle runtimeStyle
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.MSHTMLApi.IHTMLStyle>(this, "runtimeStyle", NetOffice.MSHTMLApi.IHTMLStyle.LateBindingApiWrapperType);
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

		#endregion

		#region Methods

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
		/// <param name="mergeThis">NetOffice.MSHTMLApi.IHTMLElement mergeThis</param>
		[SupportByVersion("MSHTML", 4)]
		public void mergeAttributes(NetOffice.MSHTMLApi.IHTMLElement mergeThis)
		{
			 Factory.ExecuteMethod(this, "mergeAttributes", mergeThis);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="where">string where</param>
		/// <param name="insertedElement">NetOffice.MSHTMLApi.IHTMLElement insertedElement</param>
		[SupportByVersion("MSHTML", 4)]
		public NetOffice.MSHTMLApi.IHTMLElement insertAdjacentElement(string where, NetOffice.MSHTMLApi.IHTMLElement insertedElement)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.MSHTMLApi.IHTMLElement>(this, "insertAdjacentElement", NetOffice.MSHTMLApi.IHTMLElement.LateBindingApiWrapperType, where, insertedElement);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="apply">NetOffice.MSHTMLApi.IHTMLElement apply</param>
		/// <param name="where">string where</param>
		[SupportByVersion("MSHTML", 4)]
		public NetOffice.MSHTMLApi.IHTMLElement applyElement(NetOffice.MSHTMLApi.IHTMLElement apply, string where)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.MSHTMLApi.IHTMLElement>(this, "applyElement", NetOffice.MSHTMLApi.IHTMLElement.LateBindingApiWrapperType, apply, where);
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
		public NetOffice.MSHTMLApi.IHTMLElementCollection getElementsByTagName(string v)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.MSHTMLApi.IHTMLElementCollection>(this, "getElementsByTagName", NetOffice.MSHTMLApi.IHTMLElementCollection.LateBindingApiWrapperType, v);
		}

		#endregion

		#pragma warning restore
	}
}



