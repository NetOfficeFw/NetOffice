using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.MSHTMLApi
{
	/// <summary>
	/// DispatchInterface IHTMLElement3 
	/// SupportByVersion MSHTML, 4
	/// </summary>
	[SupportByVersion("MSHTML", 4)]
	[EntityType(EntityType.IsDispatchInterface)]
 	public class IHTMLElement3 : IHTMLElement2
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
                    _type = typeof(IHTMLElement3);
                return _type;
            }
        }
        
        #endregion
        
		#region Construction

		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
		public IHTMLElement3(Core factory, ICOMObject parentObject, object comProxy) : base(factory, parentObject, comProxy)
		{
			
		}

        ///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public IHTMLElement3(ICOMObject parentObject, object comProxy) : base(parentObject, comProxy)
		{
		}
		
		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public IHTMLElement3(Core factory, ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(factory, parentObject, comProxy, comProxyType)
		{

		}

		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public IHTMLElement3(ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(parentObject, comProxy, comProxyType)
		{
		}
		
		///<param name="replacedObject">object to replaced. replacedObject are not usable after this action</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public IHTMLElement3(ICOMObject replacedObject) : base(replacedObject)
		{
		}
		
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public IHTMLElement3() : base()
		{
		}
		
		/// <param name="progId">registered progID</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public IHTMLElement3(string progId) : base(progId)
		{
		}
		
		#endregion
		
		#region Properties

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public bool isMultiLine
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "isMultiLine");
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public bool canHaveHTML
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "canHaveHTML");
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public object onlayoutcomplete
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "onlayoutcomplete");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "onlayoutcomplete", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public object onpage
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "onpage");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "onpage", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public bool inflateBlock
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "inflateBlock");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "inflateBlock", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public object onbeforedeactivate
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "onbeforedeactivate");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "onbeforedeactivate", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public string contentEditable
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "contentEditable");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "contentEditable", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public bool isContentEditable
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "isContentEditable");
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public bool hideFocus
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "hideFocus");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "hideFocus", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public bool disabled
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "disabled");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "disabled", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public bool isDisabled
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "isDisabled");
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public object onmove
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "onmove");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "onmove", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public object oncontrolselect
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "oncontrolselect");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "oncontrolselect", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public object onresizestart
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "onresizestart");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "onresizestart", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public object onresizeend
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "onresizeend");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "onresizeend", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public object onmovestart
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "onmovestart");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "onmovestart", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public object onmoveend
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "onmoveend");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "onmoveend", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public object onmouseenter
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "onmouseenter");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "onmouseenter", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public object onmouseleave
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "onmouseleave");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "onmouseleave", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public object onactivate
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "onactivate");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "onactivate", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public object ondeactivate
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "ondeactivate");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "ondeactivate", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Int32 glyphMode
		{
			get
			{
				return Factory.ExecuteInt32PropertyGet(this, "glyphMode");
			}
		}

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="mergeThis">NetOffice.MSHTMLApi.IHTMLElement mergeThis</param>
		/// <param name="pvarFlags">optional object pvarFlags</param>
		[SupportByVersion("MSHTML", 4)]
		public void mergeAttributes(NetOffice.MSHTMLApi.IHTMLElement mergeThis, object pvarFlags)
		{
			 Factory.ExecuteMethod(this, "mergeAttributes", mergeThis, pvarFlags);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="mergeThis">NetOffice.MSHTMLApi.IHTMLElement mergeThis</param>
		[CustomMethod]
		[SupportByVersion("MSHTML", 4)]
		public void mergeAttributes(NetOffice.MSHTMLApi.IHTMLElement mergeThis)
		{
			 Factory.ExecuteMethod(this, "mergeAttributes", mergeThis);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public void setActive()
		{
			 Factory.ExecuteMethod(this, "setActive");
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="bstrEventName">string bstrEventName</param>
		/// <param name="pvarEventObject">optional object pvarEventObject</param>
		[SupportByVersion("MSHTML", 4)]
		public bool FireEvent(string bstrEventName, object pvarEventObject)
		{
			return Factory.ExecuteBoolMethodGet(this, "FireEvent", bstrEventName, pvarEventObject);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="bstrEventName">string bstrEventName</param>
		[CustomMethod]
		[SupportByVersion("MSHTML", 4)]
		public bool FireEvent(string bstrEventName)
		{
			return Factory.ExecuteBoolMethodGet(this, "FireEvent", bstrEventName);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public bool dragDrop()
		{
			return Factory.ExecuteBoolMethodGet(this, "dragDrop");
		}

		#endregion

		#pragma warning restore
	}
}



