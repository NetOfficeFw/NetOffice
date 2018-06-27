using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;
using NetOffice.MSHTMLApi;

namespace NetOffice.MSHTMLApi.Behind
{
	/// <summary>
	/// DispatchInterface IHTMLAnchorElement 
	/// SupportByVersion MSHTML, 4
	/// </summary>
	[SupportByVersion("MSHTML", 4)]
	[EntityType(EntityType.IsDispatchInterface), BaseType]
 	public class IHTMLAnchorElement : COMObject, NetOffice.MSHTMLApi.IHTMLAnchorElement
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
                    _contractType = typeof(NetOffice.MSHTMLApi.IHTMLAnchorElement);
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
                    _type = typeof(IHTMLAnchorElement);
                return _type;
            }
        }
        
        #endregion
        
		#region Ctor

		/// <summary>
		/// Stub Ctor, not indented to use
		/// </summary>
		public IHTMLAnchorElement() : base()
		{

		}

		#endregion
		
		#region Properties

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public virtual string href
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "href");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "href", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public virtual string target
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "target");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "target", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public virtual string rel
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "rel");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "rel", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public virtual string rev
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "rev");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "rev", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public virtual string urn
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "urn");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "urn", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public virtual string Methods
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "Methods");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "Methods", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public virtual string name
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "name");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "name", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public virtual string host
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "host");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "host", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public virtual string hostname
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "hostname");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "hostname", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public virtual string pathname
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "pathname");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "pathname", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public virtual string port
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "port");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "port", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public virtual string protocol
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "protocol");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "protocol", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public virtual string search
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "search");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "search", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public virtual string hash
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "hash");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "hash", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public virtual object onblur
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "onblur");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteVariantPropertySet(this, "onblur", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public virtual object onfocus
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "onfocus");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteVariantPropertySet(this, "onfocus", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public virtual string accessKey
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "accessKey");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "accessKey", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public virtual string protocolLong
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "protocolLong");
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public virtual string mimeType
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "mimeType");
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public virtual string nameProp
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "nameProp");
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public virtual Int16 tabIndex
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteInt16PropertyGet(this, "tabIndex");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "tabIndex", value);
			}
		}

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public virtual void focus()
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "focus");
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public virtual void blur()
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "blur");
		}

		#endregion

		#pragma warning restore
	}
}

