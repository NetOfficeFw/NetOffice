using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.MSHTMLApi
{
	/// <summary>
	/// DispatchInterface IHTMLElement5 
	/// SupportByVersion MSHTML, 4
	/// </summary>
	[SupportByVersion("MSHTML", 4)]
	[EntityType(EntityType.IsDispatchInterface), BaseType]
 	public class IHTMLElement5 : IHTMLDatabinding
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
                    _type = typeof(IHTMLElement5);
                return _type;
            }
        }
        
        #endregion
        
		#region Ctor

		/// <param name="factory">current used factory core</param>
		/// <param name="parentObject">object there has created the proxy</param>
		/// <param name="proxyShare">proxy share instead if com proxy</param>
		public IHTMLElement5(Core factory, ICOMObject parentObject, COMProxyShare proxyShare) : base(factory, parentObject, proxyShare)
		{
		}

		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
		public IHTMLElement5(Core factory, ICOMObject parentObject, object comProxy) : base(factory, parentObject, comProxy)
		{
			
		}

        ///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public IHTMLElement5(ICOMObject parentObject, object comProxy) : base(parentObject, comProxy)
		{
		}
		
		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public IHTMLElement5(Core factory, ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(factory, parentObject, comProxy, comProxyType)
		{

		}

		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public IHTMLElement5(ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(parentObject, comProxy, comProxyType)
		{
		}
		
		///<param name="replacedObject">object to replaced. replacedObject are not usable after this action</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public IHTMLElement5(ICOMObject replacedObject) : base(replacedObject)
		{
		}
		
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public IHTMLElement5() : base()
		{
		}
		
		/// <param name="progId">registered progID</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public IHTMLElement5(string progId) : base(progId)
		{
		}
		
		#endregion
		
		#region Properties

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public string role
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "role");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "role", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public string ariaBusy
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "ariaBusy");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "ariaBusy", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public string ariaChecked
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "ariaChecked");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "ariaChecked", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public string ariaDisabled
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "ariaDisabled");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "ariaDisabled", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public string ariaExpanded
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "ariaExpanded");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "ariaExpanded", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public string ariaHaspopup
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "ariaHaspopup");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "ariaHaspopup", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public string ariaHidden
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "ariaHidden");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "ariaHidden", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public string ariaInvalid
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "ariaInvalid");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "ariaInvalid", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public string ariaMultiselectable
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "ariaMultiselectable");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "ariaMultiselectable", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public string ariaPressed
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "ariaPressed");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "ariaPressed", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public string ariaReadonly
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "ariaReadonly");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "ariaReadonly", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public string ariaRequired
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "ariaRequired");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "ariaRequired", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public string ariaSecret
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "ariaSecret");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "ariaSecret", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public string ariaSelected
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "ariaSelected");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "ariaSelected", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public NetOffice.MSHTMLApi.IHTMLAttributeCollection3 attributes
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.MSHTMLApi.IHTMLAttributeCollection3>(this, "attributes", NetOffice.MSHTMLApi.IHTMLAttributeCollection3.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public string ariaValuenow
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "ariaValuenow");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "ariaValuenow", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public Int16 ariaPosinset
		{
			get
			{
				return Factory.ExecuteInt16PropertyGet(this, "ariaPosinset");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "ariaPosinset", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public Int16 ariaSetsize
		{
			get
			{
				return Factory.ExecuteInt16PropertyGet(this, "ariaSetsize");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "ariaSetsize", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public Int16 ariaLevel
		{
			get
			{
				return Factory.ExecuteInt16PropertyGet(this, "ariaLevel");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "ariaLevel", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public string ariaValuemin
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "ariaValuemin");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "ariaValuemin", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public string ariaValuemax
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "ariaValuemax");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "ariaValuemax", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public string ariaControls
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "ariaControls");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "ariaControls", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public string ariaDescribedby
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "ariaDescribedby");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "ariaDescribedby", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public string ariaFlowto
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "ariaFlowto");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "ariaFlowto", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public string ariaLabelledby
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "ariaLabelledby");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "ariaLabelledby", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public string ariaActivedescendant
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "ariaActivedescendant");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "ariaActivedescendant", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public string ariaOwns
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "ariaOwns");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "ariaOwns", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public string ariaLive
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "ariaLive");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "ariaLive", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public string ariaRelevant
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "ariaRelevant");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "ariaRelevant", value);
			}
		}

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="bstrName">string bstrName</param>
		[SupportByVersion("MSHTML", 4)]
		[BaseResult]
		public NetOffice.MSHTMLApi.IHTMLDOMAttribute2 getAttributeNode(string bstrName)
		{
			return Factory.ExecuteBaseReferenceMethodGet<NetOffice.MSHTMLApi.IHTMLDOMAttribute2>(this, "getAttributeNode", bstrName);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="pattr">NetOffice.MSHTMLApi.IHTMLDOMAttribute2 pattr</param>
		[SupportByVersion("MSHTML", 4)]
		[BaseResult]
		public NetOffice.MSHTMLApi.IHTMLDOMAttribute2 setAttributeNode(NetOffice.MSHTMLApi.IHTMLDOMAttribute2 pattr)
		{
			return Factory.ExecuteBaseReferenceMethodGet<NetOffice.MSHTMLApi.IHTMLDOMAttribute2>(this, "setAttributeNode", pattr);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="pattr">NetOffice.MSHTMLApi.IHTMLDOMAttribute2 pattr</param>
		[SupportByVersion("MSHTML", 4)]
		[BaseResult]
		public NetOffice.MSHTMLApi.IHTMLDOMAttribute2 removeAttributeNode(NetOffice.MSHTMLApi.IHTMLDOMAttribute2 pattr)
		{
			return Factory.ExecuteBaseReferenceMethodGet<NetOffice.MSHTMLApi.IHTMLDOMAttribute2>(this, "removeAttributeNode", pattr);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="name">string name</param>
		[SupportByVersion("MSHTML", 4)]
		public bool hasAttribute(string name)
		{
			return Factory.ExecuteBoolMethodGet(this, "hasAttribute", name);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="strAttributeName">string strAttributeName</param>
		[SupportByVersion("MSHTML", 4)]
		public object getAttribute(string strAttributeName)
		{
			return Factory.ExecuteVariantMethodGet(this, "getAttribute", strAttributeName);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="strAttributeName">string strAttributeName</param>
		/// <param name="attributeValue">object attributeValue</param>
		[SupportByVersion("MSHTML", 4)]
		public void setAttribute(string strAttributeName, object attributeValue)
		{
			 Factory.ExecuteMethod(this, "setAttribute", strAttributeName, attributeValue);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="strAttributeName">string strAttributeName</param>
		[SupportByVersion("MSHTML", 4)]
		public bool removeAttribute(string strAttributeName)
		{
			return Factory.ExecuteBoolMethodGet(this, "removeAttribute", strAttributeName);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public bool hasAttributes()
		{
			return Factory.ExecuteBoolMethodGet(this, "hasAttributes");
		}

		#endregion

		#pragma warning restore
	}
}
