using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.MSHTMLApi
{
	/// <summary>
	/// DispatchInterface IHTMLAppBehavior 
	/// SupportByVersion MSHTML, 4
	/// </summary>
	[SupportByVersion("MSHTML", 4)]
	[EntityType(EntityType.IsDispatchInterface)]
 	public class IHTMLAppBehavior : DispHTMLAppBehavior
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
                    _type = typeof(IHTMLAppBehavior);
                return _type;
            }
        }
        
        #endregion
        
		#region Construction

		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
		public IHTMLAppBehavior(Core factory, ICOMObject parentObject, object comProxy) : base(factory, parentObject, comProxy)
		{
			
		}

        ///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public IHTMLAppBehavior(ICOMObject parentObject, object comProxy) : base(parentObject, comProxy)
		{
		}
		
		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public IHTMLAppBehavior(Core factory, ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(factory, parentObject, comProxy, comProxyType)
		{

		}

		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public IHTMLAppBehavior(ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(parentObject, comProxy, comProxyType)
		{
		}
		
		///<param name="replacedObject">object to replaced. replacedObject are not usable after this action</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public IHTMLAppBehavior(ICOMObject replacedObject) : base(replacedObject)
		{
		}
		
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public IHTMLAppBehavior() : base()
		{
		}
		
		/// <param name="progId">registered progID</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public IHTMLAppBehavior(string progId) : base(progId)
		{
		}
		
		#endregion
		
		#region Properties

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public string applicationName
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "applicationName");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "applicationName", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public string version
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "version");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "version", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public string icon
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "icon");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "icon", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public string singleInstance
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "singleInstance");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "singleInstance", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public string minimizeButton
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "minimizeButton");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "minimizeButton", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public string maximizeButton
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "maximizeButton");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "maximizeButton", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public string border
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "border");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "border", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public string borderStyle
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "borderStyle");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "borderStyle", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public string sysMenu
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "sysMenu");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "sysMenu", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public string caption
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "caption");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "caption", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public string windowState
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "windowState");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "windowState", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public string showInTaskBar
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "showInTaskBar");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "showInTaskBar", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public string commandLine
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "commandLine");
			}
		}

		#endregion

		#region Methods

		#endregion

		#pragma warning restore
	}
}


