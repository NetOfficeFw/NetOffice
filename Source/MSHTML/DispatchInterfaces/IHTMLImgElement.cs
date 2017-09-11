using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.MSHTMLApi
{
	/// <summary>
	/// DispatchInterface IHTMLImgElement 
	/// SupportByVersion MSHTML, 4
	/// </summary>
	[SupportByVersion("MSHTML", 4)]
	[EntityType(EntityType.IsDispatchInterface), BaseType]
 	public class IHTMLImgElement : IHTMLControlElement
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
                    _type = typeof(IHTMLImgElement);
                return _type;
            }
        }
        
        #endregion
        
		#region Ctor

		/// <param name="factory">current used factory core</param>
		/// <param name="parentObject">object there has created the proxy</param>
		/// <param name="proxyShare">proxy share instead if com proxy</param>
		public IHTMLImgElement(Core factory, ICOMObject parentObject, COMProxyShare proxyShare) : base(factory, parentObject, proxyShare)
		{
		}

		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
		public IHTMLImgElement(Core factory, ICOMObject parentObject, object comProxy) : base(factory, parentObject, comProxy)
		{
			
		}

        ///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public IHTMLImgElement(ICOMObject parentObject, object comProxy) : base(parentObject, comProxy)
		{
		}
		
		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public IHTMLImgElement(Core factory, ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(factory, parentObject, comProxy, comProxyType)
		{

		}

		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public IHTMLImgElement(ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(parentObject, comProxy, comProxyType)
		{
		}
		
		///<param name="replacedObject">object to replaced. replacedObject are not usable after this action</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public IHTMLImgElement(ICOMObject replacedObject) : base(replacedObject)
		{
		}
		
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public IHTMLImgElement() : base()
		{
		}
		
		/// <param name="progId">registered progID</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public IHTMLImgElement(string progId) : base(progId)
		{
		}
		
		#endregion
		
		#region Properties

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public bool isMap
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "isMap");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "isMap", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public string useMap
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "useMap");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "useMap", value);
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
		public string fileSize
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "fileSize");
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public string fileCreatedDate
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "fileCreatedDate");
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public string fileModifiedDate
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "fileModifiedDate");
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public string fileUpdatedDate
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "fileUpdatedDate");
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public string protocol
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "protocol");
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public string href
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "href");
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
		public object border
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "border");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "border", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public Int32 vspace
		{
			get
			{
				return Factory.ExecuteInt32PropertyGet(this, "vspace");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "vspace", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public Int32 hspace
		{
			get
			{
				return Factory.ExecuteInt32PropertyGet(this, "hspace");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "hspace", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public string alt
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "alt");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "alt", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public string src
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "src");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "src", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public string lowsrc
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "lowsrc");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "lowsrc", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public string vrml
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "vrml");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "vrml", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public string dynsrc
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "dynsrc");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "dynsrc", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public string readyState
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "readyState");
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public bool complete
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "complete");
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public object loop
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "loop");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "loop", value);
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
		public object onerror
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "onerror");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "onerror", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public object onabort
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "onabort");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "onabort", value);
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
		public Int32 width
		{
			get
			{
				return Factory.ExecuteInt32PropertyGet(this, "width");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "width", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public Int32 height
		{
			get
			{
				return Factory.ExecuteInt32PropertyGet(this, "height");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "height", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public string Start
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "Start");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "Start", value);
			}
		}

		#endregion

		#region Methods

		#endregion

		#pragma warning restore
	}
}
