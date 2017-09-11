using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.MSHTMLApi
{
	/// <summary>
	/// DispatchInterface DispHTMLRuleStyle 
	/// SupportByVersion MSHTML, 4
	/// </summary>
	[SupportByVersion("MSHTML", 4)]
	[EntityType(EntityType.IsDispatchInterface), BaseType]
 	public class DispHTMLRuleStyle : COMObject
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
                    _type = typeof(DispHTMLRuleStyle);
                return _type;
            }
        }
        
        #endregion
        
		#region Ctor

		/// <param name="factory">current used factory core</param>
		/// <param name="parentObject">object there has created the proxy</param>
		/// <param name="proxyShare">proxy share instead if com proxy</param>
		public DispHTMLRuleStyle(Core factory, ICOMObject parentObject, COMProxyShare proxyShare) : base(factory, parentObject, proxyShare)
		{
		}

		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
		public DispHTMLRuleStyle(Core factory, ICOMObject parentObject, object comProxy) : base(factory, parentObject, comProxy)
		{
			
		}

        ///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public DispHTMLRuleStyle(ICOMObject parentObject, object comProxy) : base(parentObject, comProxy)
		{
		}
		
		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public DispHTMLRuleStyle(Core factory, ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(factory, parentObject, comProxy, comProxyType)
		{

		}

		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public DispHTMLRuleStyle(ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(parentObject, comProxy, comProxyType)
		{
		}
		
		///<param name="replacedObject">object to replaced. replacedObject are not usable after this action</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public DispHTMLRuleStyle(ICOMObject replacedObject) : base(replacedObject)
		{
		}
		
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public DispHTMLRuleStyle() : base()
		{
		}
		
		/// <param name="progId">registered progID</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public DispHTMLRuleStyle(string progId) : base(progId)
		{
		}
		
		#endregion
		
		#region Properties

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public string fontFamily
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "fontFamily");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "fontFamily", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public string fontStyle
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "fontStyle");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "fontStyle", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public string fontVariant
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "fontVariant");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "fontVariant", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public string fontWeight
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "fontWeight");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "fontWeight", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public object fontSize
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "fontSize");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "fontSize", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public string font
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "font");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "font", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public object color
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "color");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "color", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public string background
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "background");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "background", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public object backgroundColor
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "backgroundColor");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "backgroundColor", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public string backgroundImage
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "backgroundImage");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "backgroundImage", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public string backgroundRepeat
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "backgroundRepeat");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "backgroundRepeat", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public string backgroundAttachment
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "backgroundAttachment");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "backgroundAttachment", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public string backgroundPosition
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "backgroundPosition");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "backgroundPosition", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public object backgroundPositionX
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "backgroundPositionX");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "backgroundPositionX", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public object backgroundPositionY
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "backgroundPositionY");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "backgroundPositionY", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public object wordSpacing
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "wordSpacing");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "wordSpacing", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public object letterSpacing
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "letterSpacing");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "letterSpacing", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public string textDecoration
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "textDecoration");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "textDecoration", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public bool textDecorationNone
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "textDecorationNone");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "textDecorationNone", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public bool textDecorationUnderline
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "textDecorationUnderline");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "textDecorationUnderline", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public bool textDecorationOverline
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "textDecorationOverline");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "textDecorationOverline", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public bool textDecorationLineThrough
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "textDecorationLineThrough");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "textDecorationLineThrough", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public bool textDecorationBlink
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "textDecorationBlink");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "textDecorationBlink", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public object verticalAlign
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "verticalAlign");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "verticalAlign", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public string textTransform
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "textTransform");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "textTransform", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public string textAlign
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "textAlign");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "textAlign", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public object textIndent
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "textIndent");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "textIndent", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public object lineHeight
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "lineHeight");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "lineHeight", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public object marginTop
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "marginTop");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "marginTop", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public object marginRight
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "marginRight");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "marginRight", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public object marginBottom
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "marginBottom");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "marginBottom", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public object marginLeft
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "marginLeft");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "marginLeft", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public string margin
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "margin");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "margin", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public object paddingTop
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "paddingTop");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "paddingTop", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public object paddingRight
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "paddingRight");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "paddingRight", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public object paddingBottom
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "paddingBottom");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "paddingBottom", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public object paddingLeft
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "paddingLeft");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "paddingLeft", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public string padding
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "padding");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "padding", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public string border
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "border");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "border", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public string borderTop
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "borderTop");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "borderTop", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public string borderRight
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "borderRight");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "borderRight", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public string borderBottom
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "borderBottom");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "borderBottom", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public string borderLeft
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "borderLeft");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "borderLeft", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public string borderColor
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "borderColor");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "borderColor", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public object borderTopColor
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "borderTopColor");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "borderTopColor", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public object borderRightColor
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "borderRightColor");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "borderRightColor", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public object borderBottomColor
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "borderBottomColor");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "borderBottomColor", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public object borderLeftColor
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "borderLeftColor");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "borderLeftColor", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public string borderWidth
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "borderWidth");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "borderWidth", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public object borderTopWidth
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "borderTopWidth");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "borderTopWidth", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public object borderRightWidth
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "borderRightWidth");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "borderRightWidth", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public object borderBottomWidth
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "borderBottomWidth");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "borderBottomWidth", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public object borderLeftWidth
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "borderLeftWidth");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "borderLeftWidth", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public string borderStyle
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "borderStyle");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "borderStyle", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public string borderTopStyle
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "borderTopStyle");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "borderTopStyle", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public string borderRightStyle
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "borderRightStyle");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "borderRightStyle", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public string borderBottomStyle
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "borderBottomStyle");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "borderBottomStyle", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public string borderLeftStyle
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "borderLeftStyle");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "borderLeftStyle", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public object width
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "width");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "width", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public object height
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "height");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "height", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public string styleFloat
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "styleFloat");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "styleFloat", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public string clear
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "clear");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "clear", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public string display
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "display");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "display", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public string visibility
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "visibility");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "visibility", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public string listStyleType
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "listStyleType");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "listStyleType", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public string listStylePosition
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "listStylePosition");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "listStylePosition", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public string listStyleImage
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "listStyleImage");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "listStyleImage", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public string listStyle
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "listStyle");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "listStyle", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public string whiteSpace
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "whiteSpace");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "whiteSpace", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public object top
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "top");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "top", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public object left
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "left");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "left", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public object zIndex
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "zIndex");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "zIndex", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public string overflow
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "overflow");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "overflow", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public string pageBreakBefore
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "pageBreakBefore");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "pageBreakBefore", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public string pageBreakAfter
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "pageBreakAfter");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "pageBreakAfter", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
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
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public string cursor
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "cursor");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "cursor", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public string clip
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "clip");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "clip", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public string filter
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "filter");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "filter", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public string tableLayout
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "tableLayout");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "tableLayout", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public string borderCollapse
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "borderCollapse");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "borderCollapse", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public string direction
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "direction");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "direction", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public string behavior
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "behavior");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "behavior", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public string position
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "position");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "position", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public string unicodeBidi
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "unicodeBidi");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "unicodeBidi", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public object bottom
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "bottom");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "bottom", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public object right
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "right");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "right", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Int32 pixelBottom
		{
			get
			{
				return Factory.ExecuteInt32PropertyGet(this, "pixelBottom");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "pixelBottom", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Int32 pixelRight
		{
			get
			{
				return Factory.ExecuteInt32PropertyGet(this, "pixelRight");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "pixelRight", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public Single posBottom
		{
			get
			{
				return Factory.ExecuteSinglePropertyGet(this, "posBottom");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "posBottom", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public Single posRight
		{
			get
			{
				return Factory.ExecuteSinglePropertyGet(this, "posRight");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "posRight", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public string imeMode
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "imeMode");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "imeMode", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public string rubyAlign
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "rubyAlign");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "rubyAlign", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public string rubyPosition
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "rubyPosition");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "rubyPosition", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public string rubyOverhang
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "rubyOverhang");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "rubyOverhang", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public object layoutGridChar
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "layoutGridChar");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "layoutGridChar", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public object layoutGridLine
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "layoutGridLine");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "layoutGridLine", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public string layoutGridMode
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "layoutGridMode");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "layoutGridMode", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public string layoutGridType
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "layoutGridType");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "layoutGridType", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public string layoutGrid
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "layoutGrid");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "layoutGrid", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public string textAutospace
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "textAutospace");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "textAutospace", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public string wordBreak
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "wordBreak");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "wordBreak", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public string lineBreak
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "lineBreak");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "lineBreak", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public string textJustify
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "textJustify");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "textJustify", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public string textJustifyTrim
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "textJustifyTrim");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "textJustifyTrim", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public object textKashida
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "textKashida");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "textKashida", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public string overflowX
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "overflowX");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "overflowX", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public string overflowY
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "overflowY");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "overflowY", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public string accelerator
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "accelerator");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "accelerator", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public string layoutFlow
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "layoutFlow");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "layoutFlow", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public object zoom
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "zoom");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "zoom", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public string wordWrap
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "wordWrap");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "wordWrap", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public string textUnderlinePosition
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "textUnderlinePosition");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "textUnderlinePosition", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public object scrollbarBaseColor
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "scrollbarBaseColor");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "scrollbarBaseColor", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public object scrollbarFaceColor
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "scrollbarFaceColor");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "scrollbarFaceColor", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public object scrollbar3dLightColor
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "scrollbar3dLightColor");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "scrollbar3dLightColor", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public object scrollbarShadowColor
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "scrollbarShadowColor");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "scrollbarShadowColor", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public object scrollbarHighlightColor
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "scrollbarHighlightColor");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "scrollbarHighlightColor", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public object scrollbarDarkShadowColor
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "scrollbarDarkShadowColor");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "scrollbarDarkShadowColor", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public object scrollbarArrowColor
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "scrollbarArrowColor");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "scrollbarArrowColor", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public object scrollbarTrackColor
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "scrollbarTrackColor");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "scrollbarTrackColor", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public string writingMode
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "writingMode");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "writingMode", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public string textAlignLast
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "textAlignLast");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "textAlignLast", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public object textKashidaSpace
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "textKashidaSpace");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "textKashidaSpace", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public string textOverflow
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "textOverflow");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "textOverflow", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public object minHeight
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "minHeight");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "minHeight", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public string msInterpolationMode
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "msInterpolationMode");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "msInterpolationMode", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public object maxHeight
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "maxHeight");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "maxHeight", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public object minWidth
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "minWidth");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "minWidth", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public object maxWidth
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "maxWidth");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "maxWidth", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public string content
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "content");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "content", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public string captionSide
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "captionSide");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "captionSide", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public string counterIncrement
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "counterIncrement");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "counterIncrement", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public string counterReset
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "counterReset");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "counterReset", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public string outline
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "outline");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "outline", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public object outlineWidth
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "outlineWidth");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "outlineWidth", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public string outlineStyle
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "outlineStyle");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "outlineStyle", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public object outlineColor
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "outlineColor");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "outlineColor", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public string boxSizing
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "boxSizing");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "boxSizing", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public string borderSpacing
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "borderSpacing");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "borderSpacing", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public object orphans
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "orphans");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "orphans", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public object widows
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "widows");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "widows", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public string pageBreakInside
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "pageBreakInside");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "pageBreakInside", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public string emptyCells
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "emptyCells");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "emptyCells", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public string msBlockProgression
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "msBlockProgression");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "msBlockProgression", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public string quotes
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "quotes");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "quotes", value);
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

		#endregion

		#pragma warning restore
	}
}
