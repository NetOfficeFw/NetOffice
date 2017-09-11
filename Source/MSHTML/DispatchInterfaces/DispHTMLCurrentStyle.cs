using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.MSHTMLApi
{
	/// <summary>
	/// DispatchInterface DispHTMLCurrentStyle 
	/// SupportByVersion MSHTML, 4
	/// </summary>
	[SupportByVersion("MSHTML", 4)]
	[EntityType(EntityType.IsDispatchInterface), BaseType]
 	public class DispHTMLCurrentStyle : COMObject
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
                    _type = typeof(DispHTMLCurrentStyle);
                return _type;
            }
        }
        
        #endregion
        
		#region Ctor

		/// <param name="factory">current used factory core</param>
		/// <param name="parentObject">object there has created the proxy</param>
		/// <param name="proxyShare">proxy share instead if com proxy</param>
		public DispHTMLCurrentStyle(Core factory, ICOMObject parentObject, COMProxyShare proxyShare) : base(factory, parentObject, proxyShare)
		{
		}

		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
		public DispHTMLCurrentStyle(Core factory, ICOMObject parentObject, object comProxy) : base(factory, parentObject, comProxy)
		{
			
		}

        ///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public DispHTMLCurrentStyle(ICOMObject parentObject, object comProxy) : base(parentObject, comProxy)
		{
		}
		
		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public DispHTMLCurrentStyle(Core factory, ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(factory, parentObject, comProxy, comProxyType)
		{

		}

		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public DispHTMLCurrentStyle(ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(parentObject, comProxy, comProxyType)
		{
		}
		
		///<param name="replacedObject">object to replaced. replacedObject are not usable after this action</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public DispHTMLCurrentStyle(ICOMObject replacedObject) : base(replacedObject)
		{
		}
		
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public DispHTMLCurrentStyle() : base()
		{
		}
		
		/// <param name="progId">registered progID</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public DispHTMLCurrentStyle(string progId) : base(progId)
		{
		}
		
		#endregion
		
		#region Properties

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public string position
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "position");
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public string styleFloat
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "styleFloat");
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public object color
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "color");
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public object backgroundColor
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "backgroundColor");
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public string fontFamily
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "fontFamily");
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public string fontStyle
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "fontStyle");
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public string fontVariant
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "fontVariant");
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public object fontWeight
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "fontWeight");
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public object fontSize
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "fontSize");
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public string backgroundImage
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "backgroundImage");
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public object backgroundPositionX
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "backgroundPositionX");
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public object backgroundPositionY
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "backgroundPositionY");
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public string backgroundRepeat
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "backgroundRepeat");
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public object borderLeftColor
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "borderLeftColor");
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public object borderTopColor
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "borderTopColor");
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public object borderRightColor
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "borderRightColor");
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public object borderBottomColor
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "borderBottomColor");
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public string borderTopStyle
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "borderTopStyle");
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public string borderRightStyle
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "borderRightStyle");
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public string borderBottomStyle
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "borderBottomStyle");
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public string borderLeftStyle
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "borderLeftStyle");
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public object borderTopWidth
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "borderTopWidth");
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public object borderRightWidth
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "borderRightWidth");
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public object borderBottomWidth
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "borderBottomWidth");
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public object borderLeftWidth
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "borderLeftWidth");
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public object left
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "left");
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public object top
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "top");
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public object width
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "width");
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public object height
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "height");
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public object paddingLeft
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "paddingLeft");
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public object paddingTop
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "paddingTop");
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public object paddingRight
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "paddingRight");
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public object paddingBottom
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "paddingBottom");
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public string textAlign
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "textAlign");
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public string textDecoration
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "textDecoration");
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public string display
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "display");
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public string visibility
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "visibility");
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public object zIndex
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "zIndex");
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public object letterSpacing
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "letterSpacing");
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public object lineHeight
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "lineHeight");
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public object textIndent
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "textIndent");
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public object verticalAlign
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "verticalAlign");
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public string backgroundAttachment
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "backgroundAttachment");
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public object marginTop
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "marginTop");
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public object marginRight
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "marginRight");
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public object marginBottom
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "marginBottom");
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public object marginLeft
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "marginLeft");
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public string clear
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "clear");
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public string listStyleType
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "listStyleType");
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public string listStylePosition
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "listStylePosition");
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public string listStyleImage
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "listStyleImage");
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public object clipTop
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "clipTop");
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public object clipRight
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "clipRight");
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public object clipBottom
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "clipBottom");
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public object clipLeft
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "clipLeft");
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public string overflow
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "overflow");
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public string pageBreakBefore
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "pageBreakBefore");
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public string pageBreakAfter
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "pageBreakAfter");
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public string cursor
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "cursor");
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public string tableLayout
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "tableLayout");
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public string borderCollapse
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "borderCollapse");
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public string direction
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "direction");
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public string behavior
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "behavior");
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public string unicodeBidi
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "unicodeBidi");
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public object right
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "right");
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public object bottom
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "bottom");
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public string imeMode
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "imeMode");
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public string rubyAlign
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "rubyAlign");
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public string rubyPosition
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "rubyPosition");
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public string rubyOverhang
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "rubyOverhang");
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public string textAutospace
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "textAutospace");
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public string lineBreak
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "lineBreak");
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public string wordBreak
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "wordBreak");
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public string textJustify
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "textJustify");
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public string textJustifyTrim
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "textJustifyTrim");
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public object textKashida
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "textKashida");
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public string blockDirection
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "blockDirection");
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public object layoutGridChar
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "layoutGridChar");
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public object layoutGridLine
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "layoutGridLine");
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public string layoutGridMode
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "layoutGridMode");
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public string layoutGridType
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "layoutGridType");
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public string borderStyle
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "borderStyle");
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public string borderColor
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "borderColor");
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public string borderWidth
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "borderWidth");
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public string padding
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "padding");
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public string margin
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "margin");
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public string accelerator
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "accelerator");
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public string overflowX
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "overflowX");
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public string overflowY
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "overflowY");
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public string textTransform
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "textTransform");
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public string layoutFlow
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "layoutFlow");
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public string wordWrap
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "wordWrap");
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public string textUnderlinePosition
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "textUnderlinePosition");
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public bool hasLayout
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "hasLayout");
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public object scrollbarBaseColor
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "scrollbarBaseColor");
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public object scrollbarFaceColor
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "scrollbarFaceColor");
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public object scrollbar3dLightColor
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "scrollbar3dLightColor");
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public object scrollbarShadowColor
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "scrollbarShadowColor");
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public object scrollbarHighlightColor
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "scrollbarHighlightColor");
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public object scrollbarDarkShadowColor
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "scrollbarDarkShadowColor");
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public object scrollbarArrowColor
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "scrollbarArrowColor");
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public object scrollbarTrackColor
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "scrollbarTrackColor");
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public string writingMode
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "writingMode");
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public object zoom
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "zoom");
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public string filter
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "filter");
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public string textAlignLast
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "textAlignLast");
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public object textKashidaSpace
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "textKashidaSpace");
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public bool isBlock
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "isBlock");
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public string textOverflow
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "textOverflow");
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public object minHeight
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "minHeight");
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public object wordSpacing
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "wordSpacing");
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public string whiteSpace
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "whiteSpace");
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public string msInterpolationMode
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "msInterpolationMode");
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public object maxHeight
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "maxHeight");
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public object minWidth
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "minWidth");
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public object maxWidth
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "maxWidth");
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public string captionSide
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "captionSide");
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public string outline
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "outline");
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public object outlineWidth
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "outlineWidth");
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public string outlineStyle
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "outlineStyle");
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public object outlineColor
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "outlineColor");
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public string boxSizing
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "boxSizing");
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public string borderSpacing
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "borderSpacing");
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public object orphans
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "orphans");
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public object widows
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "widows");
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public string pageBreakInside
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "pageBreakInside");
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public string emptyCells
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "emptyCells");
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public string msBlockProgression
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "msBlockProgression");
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public string quotes
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "quotes");
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

		#endregion

		#pragma warning restore
	}
}
