using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.MSHTMLApi
{
	/// <summary>
	/// DispatchInterface IHTMLOptionsHolder 
	/// SupportByVersion MSHTML, 4
	/// </summary>
	[SupportByVersion("MSHTML", 4)]
	[EntityType(EntityType.IsDispatchInterface)]
 	public class IHTMLOptionsHolder : COMObject
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
                    _type = typeof(IHTMLOptionsHolder);
                return _type;
            }
        }
        
        #endregion
        
		#region Ctor

		/// <param name="factory">current used factory core</param>
		/// <param name="parentObject">object there has created the proxy</param>
		/// <param name="proxyShare">proxy share instead if com proxy</param>
		public IHTMLOptionsHolder(Core factory, ICOMObject parentObject, COMProxyShare proxyShare) : base(factory, parentObject, proxyShare)
		{
		}

		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
		public IHTMLOptionsHolder(Core factory, ICOMObject parentObject, object comProxy) : base(factory, parentObject, comProxy)
		{
			
		}

        ///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public IHTMLOptionsHolder(ICOMObject parentObject, object comProxy) : base(parentObject, comProxy)
		{
		}
		
		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public IHTMLOptionsHolder(Core factory, ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(factory, parentObject, comProxy, comProxyType)
		{

		}

		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public IHTMLOptionsHolder(ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(parentObject, comProxy, comProxyType)
		{
		}
		
		///<param name="replacedObject">object to replaced. replacedObject are not usable after this action</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public IHTMLOptionsHolder(ICOMObject replacedObject) : base(replacedObject)
		{
		}
		
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public IHTMLOptionsHolder() : base()
		{
		}
		
		/// <param name="progId">registered progID</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public IHTMLOptionsHolder(string progId) : base(progId)
		{
		}
		
		#endregion
		
		#region Properties

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public NetOffice.MSHTMLApi.IHTMLDocument2 document
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.MSHTMLApi.IHTMLDocument2>(this, "document", NetOffice.MSHTMLApi.IHTMLDocument2.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public NetOffice.MSHTMLApi.IHTMLFontNamesCollection fonts
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.MSHTMLApi.IHTMLFontNamesCollection>(this, "fonts", NetOffice.MSHTMLApi.IHTMLFontNamesCollection.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public object execArg
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "execArg");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "execArg", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public Int32 errorLine
		{
			get
			{
				return Factory.ExecuteInt32PropertyGet(this, "errorLine");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "errorLine", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public Int32 errorCharacter
		{
			get
			{
				return Factory.ExecuteInt32PropertyGet(this, "errorCharacter");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "errorCharacter", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public Int32 errorCode
		{
			get
			{
				return Factory.ExecuteInt32PropertyGet(this, "errorCode");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "errorCode", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public string errorMessage
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "errorMessage");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "errorMessage", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public bool errorDebug
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "errorDebug");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "errorDebug", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public NetOffice.MSHTMLApi.IHTMLWindow2 unsecuredWindowOfDocument
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.MSHTMLApi.IHTMLWindow2>(this, "unsecuredWindowOfDocument", NetOffice.MSHTMLApi.IHTMLWindow2.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public string findText
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "findText");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "findText", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public bool anythingAfterFrameset
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "anythingAfterFrameset");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "anythingAfterFrameset", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public string secureConnectionInfo
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "secureConnectionInfo");
			}
		}

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="fontName">string fontName</param>
		[SupportByVersion("MSHTML", 4)]
		public NetOffice.MSHTMLApi.IHTMLFontSizesCollection sizes(string fontName)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.MSHTMLApi.IHTMLFontSizesCollection>(this, "sizes", NetOffice.MSHTMLApi.IHTMLFontSizesCollection.LateBindingApiWrapperType, fontName);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="initFile">optional object initFile</param>
		/// <param name="initDir">optional object initDir</param>
		/// <param name="filter">optional object filter</param>
		/// <param name="title">optional object title</param>
		[SupportByVersion("MSHTML", 4)]
		public string openfiledlg(object initFile, object initDir, object filter, object title)
		{
			return Factory.ExecuteStringMethodGet(this, "openfiledlg", initFile, initDir, filter, title);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[CustomMethod]
		[SupportByVersion("MSHTML", 4)]
		public string openfiledlg()
		{
			return Factory.ExecuteStringMethodGet(this, "openfiledlg");
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="initFile">optional object initFile</param>
		[CustomMethod]
		[SupportByVersion("MSHTML", 4)]
		public string openfiledlg(object initFile)
		{
			return Factory.ExecuteStringMethodGet(this, "openfiledlg", initFile);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="initFile">optional object initFile</param>
		/// <param name="initDir">optional object initDir</param>
		[CustomMethod]
		[SupportByVersion("MSHTML", 4)]
		public string openfiledlg(object initFile, object initDir)
		{
			return Factory.ExecuteStringMethodGet(this, "openfiledlg", initFile, initDir);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="initFile">optional object initFile</param>
		/// <param name="initDir">optional object initDir</param>
		/// <param name="filter">optional object filter</param>
		[CustomMethod]
		[SupportByVersion("MSHTML", 4)]
		public string openfiledlg(object initFile, object initDir, object filter)
		{
			return Factory.ExecuteStringMethodGet(this, "openfiledlg", initFile, initDir, filter);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="initFile">optional object initFile</param>
		/// <param name="initDir">optional object initDir</param>
		/// <param name="filter">optional object filter</param>
		/// <param name="title">optional object title</param>
		[SupportByVersion("MSHTML", 4)]
		public string savefiledlg(object initFile, object initDir, object filter, object title)
		{
			return Factory.ExecuteStringMethodGet(this, "savefiledlg", initFile, initDir, filter, title);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[CustomMethod]
		[SupportByVersion("MSHTML", 4)]
		public string savefiledlg()
		{
			return Factory.ExecuteStringMethodGet(this, "savefiledlg");
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="initFile">optional object initFile</param>
		[CustomMethod]
		[SupportByVersion("MSHTML", 4)]
		public string savefiledlg(object initFile)
		{
			return Factory.ExecuteStringMethodGet(this, "savefiledlg", initFile);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="initFile">optional object initFile</param>
		/// <param name="initDir">optional object initDir</param>
		[CustomMethod]
		[SupportByVersion("MSHTML", 4)]
		public string savefiledlg(object initFile, object initDir)
		{
			return Factory.ExecuteStringMethodGet(this, "savefiledlg", initFile, initDir);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="initFile">optional object initFile</param>
		/// <param name="initDir">optional object initDir</param>
		/// <param name="filter">optional object filter</param>
		[CustomMethod]
		[SupportByVersion("MSHTML", 4)]
		public string savefiledlg(object initFile, object initDir, object filter)
		{
			return Factory.ExecuteStringMethodGet(this, "savefiledlg", initFile, initDir, filter);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="initColor">optional object initColor</param>
		[SupportByVersion("MSHTML", 4)]
		public Int32 choosecolordlg(object initColor)
		{
			return Factory.ExecuteInt32MethodGet(this, "choosecolordlg", initColor);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[CustomMethod]
		[SupportByVersion("MSHTML", 4)]
		public Int32 choosecolordlg()
		{
			return Factory.ExecuteInt32MethodGet(this, "choosecolordlg");
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public void showSecurityInfo()
		{
			 Factory.ExecuteMethod(this, "showSecurityInfo");
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="_object">NetOffice.MSHTMLApi.IHTMLObjectElement object</param>
		[SupportByVersion("MSHTML", 4)]
		public bool isApartmentModel(NetOffice.MSHTMLApi.IHTMLObjectElement _object)
		{
			return Factory.ExecuteBoolMethodGet(this, "isApartmentModel", _object);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="fontName">string fontName</param>
		[SupportByVersion("MSHTML", 4)]
		public Int32 getCharset(string fontName)
		{
			return Factory.ExecuteInt32MethodGet(this, "getCharset", fontName);
		}

		#endregion

		#pragma warning restore
	}
}
