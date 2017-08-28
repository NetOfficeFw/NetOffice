using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.MSHTMLApi
{
	/// <summary>
	/// DispatchInterface IHTMLStyle3 
	/// SupportByVersion MSHTML, 4
	/// </summary>
	[SupportByVersion("MSHTML", 4)]
	[EntityType(EntityType.IsDispatchInterface)]
 	public class IHTMLStyle3 : IHTMLStyle2
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
                    _type = typeof(IHTMLStyle3);
                return _type;
            }
        }
        
        #endregion
        
		#region Ctor

		/// <param name="factory">current used factory core</param>
		/// <param name="parentObject">object there has created the proxy</param>
		/// <param name="proxyShare">proxy share instead if com proxy</param>
		public IHTMLStyle3(Core factory, ICOMObject parentObject, COMProxyShare proxyShare) : base(factory, parentObject, proxyShare)
		{
		}

		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
		public IHTMLStyle3(Core factory, ICOMObject parentObject, object comProxy) : base(factory, parentObject, comProxy)
		{
			
		}

        ///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public IHTMLStyle3(ICOMObject parentObject, object comProxy) : base(parentObject, comProxy)
		{
		}
		
		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public IHTMLStyle3(Core factory, ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(factory, parentObject, comProxy, comProxyType)
		{

		}

		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public IHTMLStyle3(ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(parentObject, comProxy, comProxyType)
		{
		}
		
		///<param name="replacedObject">object to replaced. replacedObject are not usable after this action</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public IHTMLStyle3(ICOMObject replacedObject) : base(replacedObject)
		{
		}
		
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public IHTMLStyle3() : base()
		{
		}
		
		/// <param name="progId">registered progID</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public IHTMLStyle3(string progId) : base(progId)
		{
		}
		
		#endregion
		
		#region Properties

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

		#endregion

		#region Methods

		#endregion

		#pragma warning restore
	}
}
