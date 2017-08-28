using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.MSHTMLApi
{
	/// <summary>
	/// DispatchInterface IHTMLRuleStyle6 
	/// SupportByVersion MSHTML, 4
	/// </summary>
	[SupportByVersion("MSHTML", 4)]
	[EntityType(EntityType.IsDispatchInterface)]
 	public class IHTMLRuleStyle6 : IHTMLRuleStyle5
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
                    _type = typeof(IHTMLRuleStyle6);
                return _type;
            }
        }
        
        #endregion
        
		#region Ctor

		/// <param name="factory">current used factory core</param>
		/// <param name="parentObject">object there has created the proxy</param>
		/// <param name="proxyShare">proxy share instead if com proxy</param>
		public IHTMLRuleStyle6(Core factory, ICOMObject parentObject, COMProxyShare proxyShare) : base(factory, parentObject, proxyShare)
		{
		}

		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
		public IHTMLRuleStyle6(Core factory, ICOMObject parentObject, object comProxy) : base(factory, parentObject, comProxy)
		{
			
		}

        ///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public IHTMLRuleStyle6(ICOMObject parentObject, object comProxy) : base(parentObject, comProxy)
		{
		}
		
		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public IHTMLRuleStyle6(Core factory, ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(factory, parentObject, comProxy, comProxyType)
		{

		}

		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public IHTMLRuleStyle6(ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(parentObject, comProxy, comProxyType)
		{
		}
		
		///<param name="replacedObject">object to replaced. replacedObject are not usable after this action</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public IHTMLRuleStyle6(ICOMObject replacedObject) : base(replacedObject)
		{
		}
		
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public IHTMLRuleStyle6() : base()
		{
		}
		
		/// <param name="progId">registered progID</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public IHTMLRuleStyle6(string progId) : base(progId)
		{
		}
		
		#endregion
		
		#region Properties

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

		#endregion

		#region Methods

		#endregion

		#pragma warning restore
	}
}
