using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.MSHTMLApi
{
	/// <summary>
	/// DispatchInterface IHTMLRuleStyle2 
	/// SupportByVersion MSHTML, 4
	/// </summary>
	[SupportByVersion("MSHTML", 4)]
	[EntityType(EntityType.IsDispatchInterface)]
 	public class IHTMLRuleStyle2 : IHTMLRuleStyle
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
                    _type = typeof(IHTMLRuleStyle2);
                return _type;
            }
        }
        
        #endregion
        
		#region Ctor

		/// <param name="factory">current used factory core</param>
		/// <param name="parentObject">object there has created the proxy</param>
		/// <param name="proxyShare">proxy share instead if com proxy</param>
		public IHTMLRuleStyle2(Core factory, ICOMObject parentObject, COMProxyShare proxyShare) : base(factory, parentObject, proxyShare)
		{
		}

		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
		public IHTMLRuleStyle2(Core factory, ICOMObject parentObject, object comProxy) : base(factory, parentObject, comProxy)
		{
			
		}

        ///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public IHTMLRuleStyle2(ICOMObject parentObject, object comProxy) : base(parentObject, comProxy)
		{
		}
		
		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public IHTMLRuleStyle2(Core factory, ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(factory, parentObject, comProxy, comProxyType)
		{

		}

		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public IHTMLRuleStyle2(ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(parentObject, comProxy, comProxyType)
		{
		}
		
		///<param name="replacedObject">object to replaced. replacedObject are not usable after this action</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public IHTMLRuleStyle2(ICOMObject replacedObject) : base(replacedObject)
		{
		}
		
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public IHTMLRuleStyle2() : base()
		{
		}
		
		/// <param name="progId">registered progID</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public IHTMLRuleStyle2(string progId) : base(progId)
		{
		}
		
		#endregion
		
		#region Properties

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

		#endregion

		#region Methods

		#endregion

		#pragma warning restore
	}
}
