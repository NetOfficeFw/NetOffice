using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;
using NetOffice.MSHTMLApi;

namespace NetOffice.MSHTMLApi.Behind
{
	/// <summary>
	/// DispatchInterface DispHTMLRuleStyle 
	/// SupportByVersion MSHTML, 4
	/// </summary>
	[SupportByVersion("MSHTML", 4)]
	[EntityType(EntityType.IsDispatchInterface), BaseType]
 	public class DispHTMLRuleStyle : COMObject, NetOffice.MSHTMLApi.DispHTMLRuleStyle
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
                    _contractType = typeof(NetOffice.MSHTMLApi.DispHTMLRuleStyle);
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
                    _type = typeof(DispHTMLRuleStyle);
                return _type;
            }
        }
        
        #endregion
        
		#region Ctor

		/// <summary>
		/// Stub Ctor, not indented to use
		/// </summary>
		public DispHTMLRuleStyle() : base()
		{

		}

		#endregion
		
		#region Properties

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public virtual string fontFamily
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "fontFamily");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "fontFamily", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public virtual string fontStyle
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "fontStyle");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "fontStyle", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public virtual string fontVariant
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "fontVariant");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "fontVariant", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public virtual string fontWeight
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "fontWeight");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "fontWeight", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public virtual object fontSize
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "fontSize");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteVariantPropertySet(this, "fontSize", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public virtual string font
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "font");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "font", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public virtual object color
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "color");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteVariantPropertySet(this, "color", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public virtual string background
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "background");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "background", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public virtual object backgroundColor
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "backgroundColor");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteVariantPropertySet(this, "backgroundColor", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public virtual string backgroundImage
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "backgroundImage");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "backgroundImage", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public virtual string backgroundRepeat
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "backgroundRepeat");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "backgroundRepeat", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public virtual string backgroundAttachment
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "backgroundAttachment");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "backgroundAttachment", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public virtual string backgroundPosition
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "backgroundPosition");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "backgroundPosition", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public virtual object backgroundPositionX
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "backgroundPositionX");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteVariantPropertySet(this, "backgroundPositionX", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public virtual object backgroundPositionY
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "backgroundPositionY");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteVariantPropertySet(this, "backgroundPositionY", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public virtual object wordSpacing
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "wordSpacing");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteVariantPropertySet(this, "wordSpacing", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public virtual object letterSpacing
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "letterSpacing");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteVariantPropertySet(this, "letterSpacing", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public virtual string textDecoration
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "textDecoration");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "textDecoration", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public virtual bool textDecorationNone
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "textDecorationNone");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "textDecorationNone", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public virtual bool textDecorationUnderline
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "textDecorationUnderline");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "textDecorationUnderline", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public virtual bool textDecorationOverline
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "textDecorationOverline");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "textDecorationOverline", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public virtual bool textDecorationLineThrough
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "textDecorationLineThrough");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "textDecorationLineThrough", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public virtual bool textDecorationBlink
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "textDecorationBlink");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "textDecorationBlink", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public virtual object verticalAlign
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "verticalAlign");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteVariantPropertySet(this, "verticalAlign", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public virtual string textTransform
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "textTransform");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "textTransform", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public virtual string textAlign
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "textAlign");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "textAlign", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public virtual object textIndent
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "textIndent");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteVariantPropertySet(this, "textIndent", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public virtual object lineHeight
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "lineHeight");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteVariantPropertySet(this, "lineHeight", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public virtual object marginTop
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "marginTop");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteVariantPropertySet(this, "marginTop", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public virtual object marginRight
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "marginRight");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteVariantPropertySet(this, "marginRight", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public virtual object marginBottom
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "marginBottom");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteVariantPropertySet(this, "marginBottom", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public virtual object marginLeft
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "marginLeft");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteVariantPropertySet(this, "marginLeft", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public virtual string margin
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "margin");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "margin", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public virtual object paddingTop
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "paddingTop");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteVariantPropertySet(this, "paddingTop", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public virtual object paddingRight
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "paddingRight");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteVariantPropertySet(this, "paddingRight", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public virtual object paddingBottom
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "paddingBottom");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteVariantPropertySet(this, "paddingBottom", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public virtual object paddingLeft
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "paddingLeft");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteVariantPropertySet(this, "paddingLeft", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public virtual string padding
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "padding");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "padding", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public virtual string border
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "border");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "border", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public virtual string borderTop
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "borderTop");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "borderTop", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public virtual string borderRight
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "borderRight");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "borderRight", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public virtual string borderBottom
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "borderBottom");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "borderBottom", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public virtual string borderLeft
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "borderLeft");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "borderLeft", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public virtual string borderColor
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "borderColor");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "borderColor", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public virtual object borderTopColor
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "borderTopColor");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteVariantPropertySet(this, "borderTopColor", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public virtual object borderRightColor
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "borderRightColor");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteVariantPropertySet(this, "borderRightColor", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public virtual object borderBottomColor
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "borderBottomColor");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteVariantPropertySet(this, "borderBottomColor", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public virtual object borderLeftColor
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "borderLeftColor");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteVariantPropertySet(this, "borderLeftColor", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public virtual string borderWidth
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "borderWidth");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "borderWidth", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public virtual object borderTopWidth
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "borderTopWidth");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteVariantPropertySet(this, "borderTopWidth", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public virtual object borderRightWidth
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "borderRightWidth");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteVariantPropertySet(this, "borderRightWidth", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public virtual object borderBottomWidth
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "borderBottomWidth");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteVariantPropertySet(this, "borderBottomWidth", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public virtual object borderLeftWidth
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "borderLeftWidth");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteVariantPropertySet(this, "borderLeftWidth", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public virtual string borderStyle
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "borderStyle");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "borderStyle", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public virtual string borderTopStyle
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "borderTopStyle");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "borderTopStyle", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public virtual string borderRightStyle
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "borderRightStyle");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "borderRightStyle", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public virtual string borderBottomStyle
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "borderBottomStyle");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "borderBottomStyle", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public virtual string borderLeftStyle
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "borderLeftStyle");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "borderLeftStyle", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public virtual object width
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "width");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteVariantPropertySet(this, "width", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public virtual object height
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "height");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteVariantPropertySet(this, "height", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public virtual string styleFloat
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "styleFloat");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "styleFloat", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public virtual string clear
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "clear");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "clear", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public virtual string display
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "display");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "display", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public virtual string visibility
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "visibility");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "visibility", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public virtual string listStyleType
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "listStyleType");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "listStyleType", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public virtual string listStylePosition
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "listStylePosition");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "listStylePosition", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public virtual string listStyleImage
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "listStyleImage");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "listStyleImage", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public virtual string listStyle
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "listStyle");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "listStyle", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public virtual string whiteSpace
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "whiteSpace");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "whiteSpace", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public virtual object top
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "top");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteVariantPropertySet(this, "top", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public virtual object left
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "left");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteVariantPropertySet(this, "left", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public virtual object zIndex
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "zIndex");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteVariantPropertySet(this, "zIndex", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public virtual string overflow
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "overflow");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "overflow", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public virtual string pageBreakBefore
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "pageBreakBefore");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "pageBreakBefore", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public virtual string pageBreakAfter
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "pageBreakAfter");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "pageBreakAfter", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public virtual string cssText
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "cssText");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "cssText", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public virtual string cursor
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "cursor");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "cursor", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public virtual string clip
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "clip");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "clip", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public virtual string filter
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "filter");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "filter", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public virtual string tableLayout
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "tableLayout");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "tableLayout", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public virtual string borderCollapse
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "borderCollapse");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "borderCollapse", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public virtual string direction
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "direction");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "direction", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public virtual string behavior
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "behavior");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "behavior", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public virtual string position
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "position");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "position", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public virtual string unicodeBidi
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "unicodeBidi");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "unicodeBidi", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public virtual object bottom
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "bottom");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteVariantPropertySet(this, "bottom", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public virtual object right
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "right");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteVariantPropertySet(this, "right", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public virtual Int32 pixelBottom
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteInt32PropertyGet(this, "pixelBottom");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "pixelBottom", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public virtual Int32 pixelRight
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteInt32PropertyGet(this, "pixelRight");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "pixelRight", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public virtual Single posBottom
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteSinglePropertyGet(this, "posBottom");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "posBottom", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public virtual Single posRight
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteSinglePropertyGet(this, "posRight");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "posRight", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public virtual string imeMode
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "imeMode");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "imeMode", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public virtual string rubyAlign
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "rubyAlign");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "rubyAlign", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public virtual string rubyPosition
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "rubyPosition");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "rubyPosition", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public virtual string rubyOverhang
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "rubyOverhang");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "rubyOverhang", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public virtual object layoutGridChar
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "layoutGridChar");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteVariantPropertySet(this, "layoutGridChar", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public virtual object layoutGridLine
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "layoutGridLine");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteVariantPropertySet(this, "layoutGridLine", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public virtual string layoutGridMode
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "layoutGridMode");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "layoutGridMode", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public virtual string layoutGridType
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "layoutGridType");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "layoutGridType", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public virtual string layoutGrid
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "layoutGrid");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "layoutGrid", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public virtual string textAutospace
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "textAutospace");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "textAutospace", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public virtual string wordBreak
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "wordBreak");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "wordBreak", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public virtual string lineBreak
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "lineBreak");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "lineBreak", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public virtual string textJustify
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "textJustify");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "textJustify", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public virtual string textJustifyTrim
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "textJustifyTrim");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "textJustifyTrim", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public virtual object textKashida
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "textKashida");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteVariantPropertySet(this, "textKashida", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public virtual string overflowX
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "overflowX");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "overflowX", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public virtual string overflowY
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "overflowY");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "overflowY", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public virtual string accelerator
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "accelerator");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "accelerator", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public virtual string layoutFlow
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "layoutFlow");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "layoutFlow", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public virtual object zoom
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "zoom");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteVariantPropertySet(this, "zoom", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public virtual string wordWrap
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "wordWrap");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "wordWrap", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public virtual string textUnderlinePosition
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "textUnderlinePosition");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "textUnderlinePosition", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public virtual object scrollbarBaseColor
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "scrollbarBaseColor");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteVariantPropertySet(this, "scrollbarBaseColor", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public virtual object scrollbarFaceColor
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "scrollbarFaceColor");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteVariantPropertySet(this, "scrollbarFaceColor", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public virtual object scrollbar3dLightColor
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "scrollbar3dLightColor");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteVariantPropertySet(this, "scrollbar3dLightColor", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public virtual object scrollbarShadowColor
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "scrollbarShadowColor");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteVariantPropertySet(this, "scrollbarShadowColor", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public virtual object scrollbarHighlightColor
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "scrollbarHighlightColor");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteVariantPropertySet(this, "scrollbarHighlightColor", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public virtual object scrollbarDarkShadowColor
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "scrollbarDarkShadowColor");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteVariantPropertySet(this, "scrollbarDarkShadowColor", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public virtual object scrollbarArrowColor
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "scrollbarArrowColor");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteVariantPropertySet(this, "scrollbarArrowColor", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public virtual object scrollbarTrackColor
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "scrollbarTrackColor");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteVariantPropertySet(this, "scrollbarTrackColor", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public virtual string writingMode
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "writingMode");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "writingMode", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public virtual string textAlignLast
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "textAlignLast");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "textAlignLast", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public virtual object textKashidaSpace
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "textKashidaSpace");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteVariantPropertySet(this, "textKashidaSpace", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public virtual string textOverflow
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "textOverflow");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "textOverflow", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public virtual object minHeight
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "minHeight");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteVariantPropertySet(this, "minHeight", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public virtual string msInterpolationMode
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "msInterpolationMode");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "msInterpolationMode", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public virtual object maxHeight
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "maxHeight");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteVariantPropertySet(this, "maxHeight", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public virtual object minWidth
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "minWidth");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteVariantPropertySet(this, "minWidth", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public virtual object maxWidth
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "maxWidth");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteVariantPropertySet(this, "maxWidth", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public virtual string content
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "content");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "content", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public virtual string captionSide
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "captionSide");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "captionSide", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public virtual string counterIncrement
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "counterIncrement");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "counterIncrement", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public virtual string counterReset
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "counterReset");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "counterReset", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public virtual string outline
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "outline");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "outline", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public virtual object outlineWidth
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "outlineWidth");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteVariantPropertySet(this, "outlineWidth", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public virtual string outlineStyle
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "outlineStyle");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "outlineStyle", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public virtual object outlineColor
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "outlineColor");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteVariantPropertySet(this, "outlineColor", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public virtual string boxSizing
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "boxSizing");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "boxSizing", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public virtual string borderSpacing
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "borderSpacing");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "borderSpacing", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public virtual object orphans
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "orphans");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteVariantPropertySet(this, "orphans", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public virtual object widows
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "widows");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteVariantPropertySet(this, "widows", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public virtual string pageBreakInside
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "pageBreakInside");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "pageBreakInside", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public virtual string emptyCells
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "emptyCells");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "emptyCells", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public virtual string msBlockProgression
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "msBlockProgression");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "msBlockProgression", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public virtual string quotes
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "quotes");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "quotes", value);
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

		#endregion

		#pragma warning restore
	}
}

