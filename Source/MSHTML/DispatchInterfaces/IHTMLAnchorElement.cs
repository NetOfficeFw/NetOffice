using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.MSHTMLApi
{
	/// <summary>
	/// DispatchInterface IHTMLAnchorElement 
	/// SupportByVersion MSHTML, 4
	/// </summary>
	[SupportByVersion("MSHTML", 4)]
	[EntityType(EntityType.IsDispatchInterface), BaseType]
 	public class IHTMLAnchorElement : COMObject
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
                    _type = typeof(IHTMLAnchorElement);
                return _type;
            }
        }
        
        #endregion
        
		#region Ctor

		/// <param name="factory">current used factory core</param>
		/// <param name="parentObject">object there has created the proxy</param>
		/// <param name="proxyShare">proxy share instead if com proxy</param>
		public IHTMLAnchorElement(Core factory, ICOMObject parentObject, COMProxyShare proxyShare) : base(factory, parentObject, proxyShare)
		{
		}

		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
		public IHTMLAnchorElement(Core factory, ICOMObject parentObject, object comProxy) : base(factory, parentObject, comProxy)
		{
			
		}

        ///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public IHTMLAnchorElement(ICOMObject parentObject, object comProxy) : base(parentObject, comProxy)
		{
		}
		
		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public IHTMLAnchorElement(Core factory, ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(factory, parentObject, comProxy, comProxyType)
		{

		}

		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public IHTMLAnchorElement(ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(parentObject, comProxy, comProxyType)
		{
		}
		
		///<param name="replacedObject">object to replaced. replacedObject are not usable after this action</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public IHTMLAnchorElement(ICOMObject replacedObject) : base(replacedObject)
		{
		}
		
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public IHTMLAnchorElement() : base()
		{
		}
		
		/// <param name="progId">registered progID</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public IHTMLAnchorElement(string progId) : base(progId)
		{
		}
		
		#endregion
		
		#region Properties

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public string href
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "href");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "href", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public string target
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "target");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "target", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public string rel
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "rel");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "rel", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public string rev
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "rev");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "rev", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public string urn
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "urn");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "urn", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public string Methods
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "Methods");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "Methods", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public string name
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "name");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "name", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public string host
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "host");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "host", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public string hostname
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "hostname");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "hostname", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public string pathname
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "pathname");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "pathname", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public string port
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "port");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "port", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public string protocol
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "protocol");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "protocol", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public string search
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "search");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "search", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public string hash
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "hash");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "hash", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public object onblur
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "onblur");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "onblur", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public object onfocus
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "onfocus");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "onfocus", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public string accessKey
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "accessKey");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "accessKey", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public string protocolLong
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "protocolLong");
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public string mimeType
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "mimeType");
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public string nameProp
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "nameProp");
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public Int16 tabIndex
		{
			get
			{
				return Factory.ExecuteInt16PropertyGet(this, "tabIndex");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "tabIndex", value);
			}
		}

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public void focus()
		{
			 Factory.ExecuteMethod(this, "focus");
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public void blur()
		{
			 Factory.ExecuteMethod(this, "blur");
		}

		#endregion

		#pragma warning restore
	}
}
