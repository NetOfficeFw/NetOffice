using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;
using NetOffice.MSHTMLApi;

namespace NetOffice.MSHTMLApi.Behind
{
	/// <summary>
	/// DispatchInterface IHTMLElement5 
	/// SupportByVersion MSHTML, 4
	/// </summary>
	[SupportByVersion("MSHTML", 4)]
	[EntityType(EntityType.IsDispatchInterface), BaseType]
 	public class IHTMLElement5 : IHTMLDatabinding, NetOffice.MSHTMLApi.IHTMLElement5
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
                    _contractType = typeof(NetOffice.MSHTMLApi.IHTMLElement5);
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
                    _type = typeof(IHTMLElement5);
                return _type;
            }
        }
        
        #endregion
        
		#region Ctor

		/// <summary>
		/// Stub Ctor, not indented to use
		/// </summary>
		public IHTMLElement5() : base()
		{

		}

		#endregion
		
		#region Properties

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public virtual string role
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "role");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "role", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public virtual string ariaBusy
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "ariaBusy");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "ariaBusy", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public virtual string ariaChecked
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "ariaChecked");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "ariaChecked", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public virtual string ariaDisabled
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "ariaDisabled");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "ariaDisabled", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public virtual string ariaExpanded
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "ariaExpanded");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "ariaExpanded", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public virtual string ariaHaspopup
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "ariaHaspopup");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "ariaHaspopup", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public virtual string ariaHidden
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "ariaHidden");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "ariaHidden", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public virtual string ariaInvalid
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "ariaInvalid");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "ariaInvalid", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public virtual string ariaMultiselectable
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "ariaMultiselectable");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "ariaMultiselectable", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public virtual string ariaPressed
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "ariaPressed");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "ariaPressed", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public virtual string ariaReadonly
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "ariaReadonly");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "ariaReadonly", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public virtual string ariaRequired
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "ariaRequired");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "ariaRequired", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public virtual string ariaSecret
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "ariaSecret");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "ariaSecret", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public virtual string ariaSelected
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "ariaSelected");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "ariaSelected", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public virtual NetOffice.MSHTMLApi.IHTMLAttributeCollection3 attributes
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.MSHTMLApi.IHTMLAttributeCollection3>(this, "attributes", typeof(NetOffice.MSHTMLApi.IHTMLAttributeCollection3));
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public virtual string ariaValuenow
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "ariaValuenow");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "ariaValuenow", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public virtual Int16 ariaPosinset
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteInt16PropertyGet(this, "ariaPosinset");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "ariaPosinset", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public virtual Int16 ariaSetsize
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteInt16PropertyGet(this, "ariaSetsize");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "ariaSetsize", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public virtual Int16 ariaLevel
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteInt16PropertyGet(this, "ariaLevel");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "ariaLevel", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public virtual string ariaValuemin
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "ariaValuemin");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "ariaValuemin", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public virtual string ariaValuemax
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "ariaValuemax");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "ariaValuemax", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public virtual string ariaControls
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "ariaControls");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "ariaControls", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public virtual string ariaDescribedby
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "ariaDescribedby");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "ariaDescribedby", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public virtual string ariaFlowto
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "ariaFlowto");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "ariaFlowto", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public virtual string ariaLabelledby
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "ariaLabelledby");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "ariaLabelledby", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public virtual string ariaActivedescendant
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "ariaActivedescendant");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "ariaActivedescendant", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public virtual string ariaOwns
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "ariaOwns");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "ariaOwns", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public virtual string ariaLive
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "ariaLive");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "ariaLive", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public virtual string ariaRelevant
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "ariaRelevant");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "ariaRelevant", value);
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
		public virtual NetOffice.MSHTMLApi.IHTMLDOMAttribute2 getAttributeNode(string bstrName)
		{
			return InvokerService.InvokeInternal.ExecuteBaseReferenceMethodGet<NetOffice.MSHTMLApi.IHTMLDOMAttribute2>(this, "getAttributeNode", bstrName);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="pattr">NetOffice.MSHTMLApi.IHTMLDOMAttribute2 pattr</param>
		[SupportByVersion("MSHTML", 4)]
		[BaseResult]
		public virtual NetOffice.MSHTMLApi.IHTMLDOMAttribute2 setAttributeNode(NetOffice.MSHTMLApi.IHTMLDOMAttribute2 pattr)
		{
			return InvokerService.InvokeInternal.ExecuteBaseReferenceMethodGet<NetOffice.MSHTMLApi.IHTMLDOMAttribute2>(this, "setAttributeNode", pattr);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="pattr">NetOffice.MSHTMLApi.IHTMLDOMAttribute2 pattr</param>
		[SupportByVersion("MSHTML", 4)]
		[BaseResult]
		public virtual NetOffice.MSHTMLApi.IHTMLDOMAttribute2 removeAttributeNode(NetOffice.MSHTMLApi.IHTMLDOMAttribute2 pattr)
		{
			return InvokerService.InvokeInternal.ExecuteBaseReferenceMethodGet<NetOffice.MSHTMLApi.IHTMLDOMAttribute2>(this, "removeAttributeNode", pattr);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="name">string name</param>
		[SupportByVersion("MSHTML", 4)]
		public virtual bool hasAttribute(string name)
		{
			return InvokerService.InvokeInternal.ExecuteBoolMethodGet(this, "hasAttribute", name);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="strAttributeName">string strAttributeName</param>
		[SupportByVersion("MSHTML", 4)]
		public virtual object getAttribute(string strAttributeName)
		{
			return InvokerService.InvokeInternal.ExecuteVariantMethodGet(this, "getAttribute", strAttributeName);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="strAttributeName">string strAttributeName</param>
		/// <param name="attributeValue">object attributeValue</param>
		[SupportByVersion("MSHTML", 4)]
		public virtual void setAttribute(string strAttributeName, object attributeValue)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "setAttribute", strAttributeName, attributeValue);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="strAttributeName">string strAttributeName</param>
		[SupportByVersion("MSHTML", 4)]
		public virtual bool removeAttribute(string strAttributeName)
		{
			return InvokerService.InvokeInternal.ExecuteBoolMethodGet(this, "removeAttribute", strAttributeName);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public virtual bool hasAttributes()
		{
			return InvokerService.InvokeInternal.ExecuteBoolMethodGet(this, "hasAttributes");
		}

		#endregion

		#pragma warning restore
	}
}


