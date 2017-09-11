using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.MSHTMLApi
{
	/// <summary>
	/// DispatchInterface IHTMLBodyElement 
	/// SupportByVersion MSHTML, 4
	/// </summary>
	[SupportByVersion("MSHTML", 4)]
	[EntityType(EntityType.IsDispatchInterface), BaseType]
 	public class IHTMLBodyElement : IHTMLTextContainer
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
                    _type = typeof(IHTMLBodyElement);
                return _type;
            }
        }
        
        #endregion
        
		#region Ctor

		/// <param name="factory">current used factory core</param>
		/// <param name="parentObject">object there has created the proxy</param>
		/// <param name="proxyShare">proxy share instead if com proxy</param>
		public IHTMLBodyElement(Core factory, ICOMObject parentObject, COMProxyShare proxyShare) : base(factory, parentObject, proxyShare)
		{
		}

		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
		public IHTMLBodyElement(Core factory, ICOMObject parentObject, object comProxy) : base(factory, parentObject, comProxy)
		{
			
		}

        ///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public IHTMLBodyElement(ICOMObject parentObject, object comProxy) : base(parentObject, comProxy)
		{
		}
		
		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public IHTMLBodyElement(Core factory, ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(factory, parentObject, comProxy, comProxyType)
		{

		}

		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public IHTMLBodyElement(ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(parentObject, comProxy, comProxyType)
		{
		}
		
		///<param name="replacedObject">object to replaced. replacedObject are not usable after this action</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public IHTMLBodyElement(ICOMObject replacedObject) : base(replacedObject)
		{
		}
		
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public IHTMLBodyElement() : base()
		{
		}
		
		/// <param name="progId">registered progID</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public IHTMLBodyElement(string progId) : base(progId)
		{
		}
		
		#endregion
		
		#region Properties

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
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
		public string bgProperties
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "bgProperties");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "bgProperties", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public object leftMargin
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "leftMargin");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "leftMargin", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public object topMargin
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "topMargin");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "topMargin", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public object rightMargin
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "rightMargin");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "rightMargin", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public object bottomMargin
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "bottomMargin");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "bottomMargin", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public object bgColor
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "bgColor");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "bgColor", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public object text
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "text");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "text", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public object link
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "link");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "link", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public object vLink
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "vLink");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "vLink", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public object aLink
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "aLink");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "aLink", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public object onload
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "onload");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "onload", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public object onunload
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "onunload");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "onunload", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public string scroll
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "scroll");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "scroll", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public object onselect
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "onselect");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "onselect", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public object onbeforeunload
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "onbeforeunload");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "onbeforeunload", value);
			}
		}

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public NetOffice.MSHTMLApi.IHTMLTxtRange createTextRange()
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.MSHTMLApi.IHTMLTxtRange>(this, "createTextRange", NetOffice.MSHTMLApi.IHTMLTxtRange.LateBindingApiWrapperType);
		}

		#endregion

		#pragma warning restore
	}
}
