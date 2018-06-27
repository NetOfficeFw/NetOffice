using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;
using NetOffice.MSHTMLApi;

namespace NetOffice.MSHTMLApi.Behind
{
	/// <summary>
	/// DispatchInterface IHTMLOptionsHolder 
	/// SupportByVersion MSHTML, 4
	/// </summary>
	[SupportByVersion("MSHTML", 4)]
	[EntityType(EntityType.IsDispatchInterface)]
 	public class IHTMLOptionsHolder : COMObject, NetOffice.MSHTMLApi.IHTMLOptionsHolder
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
                    _contractType = typeof(NetOffice.MSHTMLApi.IHTMLOptionsHolder);
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
                    _type = typeof(IHTMLOptionsHolder);
                return _type;
            }
        }
        
        #endregion
        
		#region Ctor

		/// <summary>
		/// Stub Ctor, not indented to use
		/// </summary>
		public IHTMLOptionsHolder() : base()
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
		public virtual NetOffice.MSHTMLApi.IHTMLDocument2 document
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.MSHTMLApi.IHTMLDocument2>(this, "document", typeof(NetOffice.MSHTMLApi.IHTMLDocument2));
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public virtual NetOffice.MSHTMLApi.IHTMLFontNamesCollection fonts
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.MSHTMLApi.IHTMLFontNamesCollection>(this, "fonts", typeof(NetOffice.MSHTMLApi.IHTMLFontNamesCollection));
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public virtual object execArg
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "execArg");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteVariantPropertySet(this, "execArg", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public virtual Int32 errorLine
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteInt32PropertyGet(this, "errorLine");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "errorLine", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public virtual Int32 errorCharacter
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteInt32PropertyGet(this, "errorCharacter");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "errorCharacter", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public virtual Int32 errorCode
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteInt32PropertyGet(this, "errorCode");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "errorCode", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public virtual string errorMessage
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "errorMessage");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "errorMessage", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public virtual bool errorDebug
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "errorDebug");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "errorDebug", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public virtual NetOffice.MSHTMLApi.IHTMLWindow2 unsecuredWindowOfDocument
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.MSHTMLApi.IHTMLWindow2>(this, "unsecuredWindowOfDocument", typeof(NetOffice.MSHTMLApi.IHTMLWindow2));
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public virtual string findText
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "findText");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "findText", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public virtual bool anythingAfterFrameset
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "anythingAfterFrameset");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "anythingAfterFrameset", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public virtual string secureConnectionInfo
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "secureConnectionInfo");
			}
		}

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="fontName">string fontName</param>
		[SupportByVersion("MSHTML", 4)]
		public virtual NetOffice.MSHTMLApi.IHTMLFontSizesCollection sizes(string fontName)
		{
			return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.MSHTMLApi.IHTMLFontSizesCollection>(this, "sizes", typeof(NetOffice.MSHTMLApi.IHTMLFontSizesCollection), fontName);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="initFile">optional object initFile</param>
		/// <param name="initDir">optional object initDir</param>
		/// <param name="filter">optional object filter</param>
		/// <param name="title">optional object title</param>
		[SupportByVersion("MSHTML", 4)]
		public virtual string openfiledlg(object initFile, object initDir, object filter, object title)
		{
			return InvokerService.InvokeInternal.ExecuteStringMethodGet(this, "openfiledlg", initFile, initDir, filter, title);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[CustomMethod]
		[SupportByVersion("MSHTML", 4)]
		public virtual string openfiledlg()
		{
			return InvokerService.InvokeInternal.ExecuteStringMethodGet(this, "openfiledlg");
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="initFile">optional object initFile</param>
		[CustomMethod]
		[SupportByVersion("MSHTML", 4)]
		public virtual string openfiledlg(object initFile)
		{
			return InvokerService.InvokeInternal.ExecuteStringMethodGet(this, "openfiledlg", initFile);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="initFile">optional object initFile</param>
		/// <param name="initDir">optional object initDir</param>
		[CustomMethod]
		[SupportByVersion("MSHTML", 4)]
		public virtual string openfiledlg(object initFile, object initDir)
		{
			return InvokerService.InvokeInternal.ExecuteStringMethodGet(this, "openfiledlg", initFile, initDir);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="initFile">optional object initFile</param>
		/// <param name="initDir">optional object initDir</param>
		/// <param name="filter">optional object filter</param>
		[CustomMethod]
		[SupportByVersion("MSHTML", 4)]
		public virtual string openfiledlg(object initFile, object initDir, object filter)
		{
			return InvokerService.InvokeInternal.ExecuteStringMethodGet(this, "openfiledlg", initFile, initDir, filter);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="initFile">optional object initFile</param>
		/// <param name="initDir">optional object initDir</param>
		/// <param name="filter">optional object filter</param>
		/// <param name="title">optional object title</param>
		[SupportByVersion("MSHTML", 4)]
		public virtual string savefiledlg(object initFile, object initDir, object filter, object title)
		{
			return InvokerService.InvokeInternal.ExecuteStringMethodGet(this, "savefiledlg", initFile, initDir, filter, title);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[CustomMethod]
		[SupportByVersion("MSHTML", 4)]
		public virtual string savefiledlg()
		{
			return InvokerService.InvokeInternal.ExecuteStringMethodGet(this, "savefiledlg");
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="initFile">optional object initFile</param>
		[CustomMethod]
		[SupportByVersion("MSHTML", 4)]
		public virtual string savefiledlg(object initFile)
		{
			return InvokerService.InvokeInternal.ExecuteStringMethodGet(this, "savefiledlg", initFile);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="initFile">optional object initFile</param>
		/// <param name="initDir">optional object initDir</param>
		[CustomMethod]
		[SupportByVersion("MSHTML", 4)]
		public virtual string savefiledlg(object initFile, object initDir)
		{
			return InvokerService.InvokeInternal.ExecuteStringMethodGet(this, "savefiledlg", initFile, initDir);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="initFile">optional object initFile</param>
		/// <param name="initDir">optional object initDir</param>
		/// <param name="filter">optional object filter</param>
		[CustomMethod]
		[SupportByVersion("MSHTML", 4)]
		public virtual string savefiledlg(object initFile, object initDir, object filter)
		{
			return InvokerService.InvokeInternal.ExecuteStringMethodGet(this, "savefiledlg", initFile, initDir, filter);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="initColor">optional object initColor</param>
		[SupportByVersion("MSHTML", 4)]
		public virtual Int32 choosecolordlg(object initColor)
		{
			return InvokerService.InvokeInternal.ExecuteInt32MethodGet(this, "choosecolordlg", initColor);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[CustomMethod]
		[SupportByVersion("MSHTML", 4)]
		public virtual Int32 choosecolordlg()
		{
			return InvokerService.InvokeInternal.ExecuteInt32MethodGet(this, "choosecolordlg");
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public virtual void showSecurityInfo()
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "showSecurityInfo");
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="_object">NetOffice.MSHTMLApi.IHTMLObjectElement object</param>
		[SupportByVersion("MSHTML", 4)]
		public virtual bool isApartmentModel(NetOffice.MSHTMLApi.IHTMLObjectElement _object)
		{
			return InvokerService.InvokeInternal.ExecuteBoolMethodGet(this, "isApartmentModel", _object);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="fontName">string fontName</param>
		[SupportByVersion("MSHTML", 4)]
		public virtual Int32 getCharset(string fontName)
		{
			return InvokerService.InvokeInternal.ExecuteInt32MethodGet(this, "getCharset", fontName);
		}

		#endregion

		#pragma warning restore
	}
}


