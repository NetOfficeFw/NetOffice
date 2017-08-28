using System.Reflection;
using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.OfficeApi
{
	/// <summary>
	/// CommandBar
	/// </summary>
	[SyntaxBypass]
 	public class CommandBar_ : _IMsoOleAccDispObj
	{
		#region Ctor

		/// <param name="factory">current used factory core</param>
		/// <param name="parentObject">object there has created the proxy</param>
		/// <param name="proxyShare">proxy share instead if com proxy</param>
		public CommandBar_(Core factory, ICOMObject parentObject, COMProxyShare proxyShare) : base(factory, parentObject, proxyShare)
		{
		}

		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
		public CommandBar_(Core factory, ICOMObject parentObject, object comProxy) : base(factory, parentObject, comProxy)
		{
			
		}

        /// <param name="parentObject">object there has created the proxy</param>
        /// <param name="comProxy">inner wrapped COM proxy</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public CommandBar_(ICOMObject parentObject, object comProxy) : base(parentObject, comProxy)
		{
		}
		
		/// <param name="parentObject">object there has created the proxy</param>
        /// <param name="comProxy">inner wrapped COM proxy</param>
        /// <param name="comProxyType">Type of inner wrapped COM proxy"</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public CommandBar_(ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(parentObject, comProxy, comProxyType)
		{
		}

		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public CommandBar_(Core factory, ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(factory, parentObject, comProxy, comProxyType)
		{

		}
		
		/// <param name="replacedObject">object to replaced. replacedObject are not usable after this action</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public CommandBar_(ICOMObject replacedObject) : base(replacedObject)
		{
		}

		/// <summary>
        /// Hidden stub .ctor
        /// </summary>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public CommandBar_() : base()
		{
		}
		
		/// <param name="progId">registered progID</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public CommandBar_(string progId) : base(progId)
		{
		}
		
		#endregion

		#region Properties

		/// <summary>
		/// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <param name="varChild">optional object varChild</param>
		[SupportByVersion("Office", 9,10,11,12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public new string get_accName(object varChild)
		{
			return Factory.ExecuteStringPropertyGet(this, "accName", varChild);
		}

        /// <summary>
        /// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <param name="varChild">optional object varChild</param>
        /// <param name="value">optional string value</param>
        [SupportByVersion("Office", 9,10,11,12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public new void set_accName(object varChild, string value)
		{
			Factory.ExecutePropertySet(this, "accName", varChild, value);
		}

		/// <summary>
		/// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
		/// Alias for get_accName
		/// </summary>
		/// <param name="varChild">optional object varChild</param>
		[SupportByVersion("Office", 9,10,11,12,14,15,16), Redirect("get_accName")]
		public new string accName(object varChild)
		{
			return get_accName(varChild);
		}

		/// <summary>
		/// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <param name="varChild">optional object varChild</param>
		[SupportByVersion("Office", 9,10,11,12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public new string get_accValue(object varChild)
		{
			return Factory.ExecuteStringPropertyGet(this, "accValue", varChild);
		}

        /// <summary>
        /// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <param name="varChild">optional object varChild</param>
        /// <param name="value">optional string value</param>
        [SupportByVersion("Office", 9,10,11,12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public new void set_accValue(object varChild, string value)
		{
			Factory.ExecutePropertySet(this, "accValue", varChild, value);
		}

		/// <summary>
		/// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
		/// Alias for get_accValue
		/// </summary>
		/// <param name="varChild">optional object varChild</param>
		[SupportByVersion("Office", 9,10,11,12,14,15,16), Redirect("get_accValue")]
		public new string accValue(object varChild)
		{
			return get_accValue(varChild);
		}

		/// <summary>
		/// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <param name="varChild">optional object varChild</param>
		[SupportByVersion("Office", 9,10,11,12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public new string get_accDescription(object varChild)
		{
			return Factory.ExecuteStringPropertyGet(this, "accDescription", varChild);
		}

		/// <summary>
		/// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
		/// Alias for get_accDescription
		/// </summary>
		/// <param name="varChild">optional object varChild</param>
		[SupportByVersion("Office", 9,10,11,12,14,15,16), Redirect("get_accDescription")]
		public new string accDescription(object varChild)
		{
			return get_accDescription(varChild);
		}

		/// <summary>
		/// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <param name="varChild">optional object varChild</param>
		[SupportByVersion("Office", 9,10,11,12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public new object get_accRole(object varChild)
		{
			return Factory.ExecuteVariantPropertyGet(this, "accRole", varChild);
		}

		/// <summary>
		/// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
		/// Alias for get_accRole
		/// </summary>
		/// <param name="varChild">optional object varChild</param>
		[SupportByVersion("Office", 9,10,11,12,14,15,16), Redirect("get_accRole")]
		public new object accRole(object varChild)
		{
			return get_accRole(varChild);
		}

		/// <summary>
		/// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <param name="varChild">optional object varChild</param>
		[SupportByVersion("Office", 9,10,11,12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public new object get_accState(object varChild)
		{
			return Factory.ExecuteVariantPropertyGet(this, "accState", varChild);
		}

		/// <summary>
		/// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
		/// Alias for get_accState
		/// </summary>
		/// <param name="varChild">optional object varChild</param>
		[SupportByVersion("Office", 9,10,11,12,14,15,16), Redirect("get_accState")]
		public new object accState(object varChild)
		{
			return get_accState(varChild);
		}

		/// <summary>
		/// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <param name="varChild">optional object varChild</param>
		[SupportByVersion("Office", 9,10,11,12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public new string get_accHelp(object varChild)
		{
			return Factory.ExecuteStringPropertyGet(this, "accHelp", varChild);
		}

		/// <summary>
		/// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
		/// Alias for get_accHelp
		/// </summary>
		/// <param name="varChild">optional object varChild</param>
		[SupportByVersion("Office", 9,10,11,12,14,15,16), Redirect("get_accHelp")]
		public new string accHelp(object varChild)
		{
			return get_accHelp(varChild);
		}

		/// <summary>
		/// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <param name="varChild">optional object varChild</param>
		[SupportByVersion("Office", 9,10,11,12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public new string get_accKeyboardShortcut(object varChild)
		{
			return Factory.ExecuteStringPropertyGet(this, "accKeyboardShortcut", varChild);
		}

		/// <summary>
		/// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
		/// Alias for get_accKeyboardShortcut
		/// </summary>
		/// <param name="varChild">optional object varChild</param>
		[SupportByVersion("Office", 9,10,11,12,14,15,16), Redirect("get_accKeyboardShortcut")]
		public new string accKeyboardShortcut(object varChild)
		{
			return get_accKeyboardShortcut(varChild);
		}

		/// <summary>
		/// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <param name="varChild">optional object varChild</param>
		[SupportByVersion("Office", 9,10,11,12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public new string get_accDefaultAction(object varChild)
		{
			return Factory.ExecuteStringPropertyGet(this, "accDefaultAction", varChild);
		}

		/// <summary>
		/// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
		/// Alias for get_accDefaultAction
		/// </summary>
		/// <param name="varChild">optional object varChild</param>
		[SupportByVersion("Office", 9,10,11,12,14,15,16), Redirect("get_accDefaultAction")]
		public new string accDefaultAction(object varChild)
		{
			return get_accDefaultAction(varChild);
		}

		#endregion

		#region Methods

		#endregion

	}

	/// <summary>
	/// DispatchInterface CommandBar 
	/// SupportByVersion Office, 9,10,11,12,14,15,16
	/// </summary>
	/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff862548.aspx </remarks>
	[SupportByVersion("Office", 9,10,11,12,14,15,16)]
	[EntityType(EntityType.IsDispatchInterface)]
 	public class CommandBar : CommandBar_
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
                    _type = typeof(CommandBar);
                return _type;
            }
        }
        
        #endregion
        
		#region Ctor

		/// <param name="factory">current used factory core</param>
		/// <param name="parentObject">object there has created the proxy</param>
		/// <param name="proxyShare">proxy share instead if com proxy</param>
		public CommandBar(Core factory, ICOMObject parentObject, COMProxyShare proxyShare) : base(factory, parentObject, proxyShare)
		{
		}

		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
		public CommandBar(Core factory, ICOMObject parentObject, object comProxy) : base(factory, parentObject, comProxy)
		{
			
		}

        ///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public CommandBar(ICOMObject parentObject, object comProxy) : base(parentObject, comProxy)
		{
		}
		
		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public CommandBar(Core factory, ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(factory, parentObject, comProxy, comProxyType)
		{

		}

		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public CommandBar(ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(parentObject, comProxy, comProxyType)
		{
		}
		
		///<param name="replacedObject">object to replaced. replacedObject are not usable after this action</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public CommandBar(ICOMObject replacedObject) : base(replacedObject)
		{
		}
		
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public CommandBar() : base()
		{
		}
		
		/// <param name="progId">registered progID</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public CommandBar(string progId) : base(progId)
		{
		}
		
		#endregion
		
		#region Properties

		/// <summary>
		/// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		[SupportByVersion("Office", 9,10,11,12,14,15,16), ProxyResult]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public object accParent
		{
			get
			{
				return Factory.ExecuteReferencePropertyGet(this, "accParent");
			}
		}

		/// <summary>
		/// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Office", 9,10,11,12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Int32 accChildCount
		{
			get
			{
				return Factory.ExecuteInt32PropertyGet(this, "accChildCount");
			}
		}

		/// <summary>
		/// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		/// <param name="varChild">object varChild</param>
		[SupportByVersion("Office", 9,10,11,12,14,15,16), ProxyResult]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public object get_accChild(object varChild)
		{
			return Factory.ExecuteReferencePropertyGet(this, "accChild", varChild);
		}

		/// <summary>
		/// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
		/// Alias for get_accChild
		/// Unknown COM Proxy
		/// </summary>
		/// <param name="varChild">object varChild</param>
		[SupportByVersion("Office", 9,10,11,12,14,15,16), ProxyResult, Redirect("get_accChild")]
		public object accChild(object varChild)
		{
			return get_accChild(varChild);
		}

		/// <summary>
		/// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Office", 9,10,11,12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public string accName
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "accName");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "accName", value);
			}
		}

		/// <summary>
		/// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Office", 9,10,11,12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public string accValue
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "accValue");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "accValue", value);
			}
		}

		/// <summary>
		/// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Office", 9,10,11,12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public string accDescription
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "accDescription");
			}
		}

		/// <summary>
		/// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Office", 9,10,11,12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public object accRole
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "accRole");
			}
		}

		/// <summary>
		/// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Office", 9,10,11,12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public object accState
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "accState");
			}
		}

		/// <summary>
		/// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Office", 9,10,11,12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public string accHelp
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "accHelp");
			}
		}

		/// <summary>
		/// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <param name="pszHelpFile">string pszHelpFile</param>
		/// <param name="varChild">optional object varChild</param>
		[SupportByVersion("Office", 9,10,11,12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Int32 get_accHelpTopic(out string pszHelpFile, object varChild)
		{		
			ParameterModifier[] modifiers = Invoker.CreateParamModifiers(true,false);
			pszHelpFile = string.Empty;
			object[] paramsArray = Invoker.ValidateParamsArray(pszHelpFile, varChild);
			object returnItem = Invoker.PropertyGet(this, "accHelpTopic", paramsArray, modifiers);
			pszHelpFile = (string)paramsArray[0];
			return NetRuntimeSystem.Convert.ToInt32(returnItem);
		}

		/// <summary>
		/// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
		/// Alias for get_accHelpTopic
		/// </summary>
		/// <param name="pszHelpFile">string pszHelpFile</param>
		/// <param name="varChild">optional object varChild</param>
		[SupportByVersion("Office", 9,10,11,12,14,15,16), Redirect("get_accHelpTopic")]
		public Int32 accHelpTopic(out string pszHelpFile, object varChild)
		{
			return get_accHelpTopic(out pszHelpFile, varChild);
		}

		/// <summary>
		/// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <param name="pszHelpFile">string pszHelpFile</param>
		[SupportByVersion("Office", 9,10,11,12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Int32 get_accHelpTopic(out string pszHelpFile)
		{		
			ParameterModifier[] modifiers = Invoker.CreateParamModifiers(true);
			pszHelpFile = string.Empty;
			object[] paramsArray = Invoker.ValidateParamsArray(pszHelpFile);
			object returnItem = Invoker.PropertyGet(this, "accHelpTopic", paramsArray, modifiers);
			pszHelpFile = (string)paramsArray[0];
			return NetRuntimeSystem.Convert.ToInt32(returnItem);
		}

		/// <summary>
		/// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
		/// Alias for get_accHelpTopic
		/// </summary>
		/// <param name="pszHelpFile">string pszHelpFile</param>
		[SupportByVersion("Office", 9,10,11,12,14,15,16), Redirect("get_accHelpTopic")]
		public Int32 accHelpTopic(out string pszHelpFile)
		{
			return get_accHelpTopic(out pszHelpFile);
		}

		/// <summary>
		/// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Office", 9,10,11,12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public string accKeyboardShortcut
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "accKeyboardShortcut");
			}
		}

		/// <summary>
		/// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Office", 9,10,11,12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public object accFocus
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "accFocus");
			}
		}

		/// <summary>
		/// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Office", 9,10,11,12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public object accSelection
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "accSelection");
			}
		}

		/// <summary>
		/// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Office", 9,10,11,12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public string accDefaultAction
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "accDefaultAction");
			}
		}

		/// <summary>
		/// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff865497.aspx </remarks>
		[SupportByVersion("Office", 9,10,11,12,14,15,16)]
		public bool BuiltIn
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "BuiltIn");
			}
		}

		/// <summary>
		/// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff865230.aspx </remarks>
		[SupportByVersion("Office", 9,10,11,12,14,15,16)]
		public string Context
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "Context");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "Context", value);
			}
		}

		/// <summary>
		/// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff861889.aspx </remarks>
		[SupportByVersion("Office", 9,10,11,12,14,15,16)]
		public NetOffice.OfficeApi.CommandBarControls Controls
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.OfficeApi.CommandBarControls>(this, "Controls", NetOffice.OfficeApi.CommandBarControls.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff861500.aspx </remarks>
		[SupportByVersion("Office", 9,10,11,12,14,15,16)]
		public bool Enabled
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "Enabled");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "Enabled", value);
			}
		}

		/// <summary>
		/// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff863298.aspx </remarks>
		[SupportByVersion("Office", 9,10,11,12,14,15,16)]
		public Int32 Height
		{
			get
			{
				return Factory.ExecuteInt32PropertyGet(this, "Height");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "Height", value);
			}
		}

		/// <summary>
		/// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff863643.aspx </remarks>
		[SupportByVersion("Office", 9,10,11,12,14,15,16)]
		public Int32 Index
		{
			get
			{
				return Factory.ExecuteInt32PropertyGet(this, "Index");
			}
		}

		/// <summary>
		/// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Office", 9,10,11,12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Int32 InstanceId
		{
			get
			{
				return Factory.ExecuteInt32PropertyGet(this, "InstanceId");
			}
		}

		/// <summary>
		/// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff860792.aspx </remarks>
		[SupportByVersion("Office", 9,10,11,12,14,15,16)]
		public Int32 Left
		{
			get
			{
				return Factory.ExecuteInt32PropertyGet(this, "Left");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "Left", value);
			}
		}

		/// <summary>
		/// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff861533.aspx </remarks>
		[SupportByVersion("Office", 9,10,11,12,14,15,16)]
		public string Name
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "Name");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "Name", value);
			}
		}

		/// <summary>
		/// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff861194.aspx </remarks>
		[SupportByVersion("Office", 9,10,11,12,14,15,16)]
		public string NameLocal
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "NameLocal");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "NameLocal", value);
			}
		}

		/// <summary>
		/// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff862362.aspx </remarks>
		[SupportByVersion("Office", 9,10,11,12,14,15,16), ProxyResult]
		public object Parent
		{
			get
			{
				return Factory.ExecuteReferencePropertyGet(this, "Parent");
			}
		}

		/// <summary>
		/// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff863844.aspx </remarks>
		[SupportByVersion("Office", 9,10,11,12,14,15,16)]
		public NetOffice.OfficeApi.Enums.MsoBarPosition Position
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.OfficeApi.Enums.MsoBarPosition>(this, "Position");
			}
			set
			{
				Factory.ExecuteEnumPropertySet(this, "Position", value);
			}
		}

		/// <summary>
		/// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff862402.aspx </remarks>
		[SupportByVersion("Office", 9,10,11,12,14,15,16)]
		public Int32 RowIndex
		{
			get
			{
				return Factory.ExecuteInt32PropertyGet(this, "RowIndex");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "RowIndex", value);
			}
		}

		/// <summary>
		/// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff861854.aspx </remarks>
		[SupportByVersion("Office", 9,10,11,12,14,15,16)]
		public NetOffice.OfficeApi.Enums.MsoBarProtection Protection
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.OfficeApi.Enums.MsoBarProtection>(this, "Protection");
			}
			set
			{
				Factory.ExecuteEnumPropertySet(this, "Protection", value);
			}
		}

		/// <summary>
		/// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff860591.aspx </remarks>
		[SupportByVersion("Office", 9,10,11,12,14,15,16)]
		public Int32 Top
		{
			get
			{
				return Factory.ExecuteInt32PropertyGet(this, "Top");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "Top", value);
			}
		}

		/// <summary>
		/// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff864969.aspx </remarks>
		[SupportByVersion("Office", 9,10,11,12,14,15,16)]
		public NetOffice.OfficeApi.Enums.MsoBarType Type
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.OfficeApi.Enums.MsoBarType>(this, "Type");
			}
		}

		/// <summary>
		/// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff864581.aspx </remarks>
		[SupportByVersion("Office", 9,10,11,12,14,15,16)]
		public bool Visible
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "Visible");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "Visible", value);
			}
		}

		/// <summary>
		/// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff863766.aspx </remarks>
		[SupportByVersion("Office", 9,10,11,12,14,15,16)]
		public Int32 Width
		{
			get
			{
				return Factory.ExecuteInt32PropertyGet(this, "Width");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "Width", value);
			}
		}

		/// <summary>
		/// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff860615.aspx </remarks>
		[SupportByVersion("Office", 9,10,11,12,14,15,16)]
		public bool AdaptiveMenu
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "AdaptiveMenu");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "AdaptiveMenu", value);
			}
		}

		/// <summary>
		/// SupportByVersion Office 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Office", 10,11,12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Int32 Id
		{
			get
			{
				return Factory.ExecuteInt32PropertyGet(this, "Id");
			}
		}

		/// <summary>
		/// SupportByVersion Office 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Office", 14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public object InstanceIdPtr
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "InstanceIdPtr");
			}
		}

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="flagsSelect">Int32 flagsSelect</param>
		/// <param name="varChild">optional object varChild</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("Office", 9,10,11,12,14,15,16)]
		public void accSelect(Int32 flagsSelect, object varChild)
		{
			 Factory.ExecuteMethod(this, "accSelect", flagsSelect, varChild);
		}

		/// <summary>
		/// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="flagsSelect">Int32 flagsSelect</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Office", 9,10,11,12,14,15,16)]
		public void accSelect(Int32 flagsSelect)
		{
			 Factory.ExecuteMethod(this, "accSelect", flagsSelect);
		}

		/// <summary>
		/// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="pxLeft">Int32 pxLeft</param>
		/// <param name="pyTop">Int32 pyTop</param>
		/// <param name="pcxWidth">Int32 pcxWidth</param>
		/// <param name="pcyHeight">Int32 pcyHeight</param>
		/// <param name="varChild">optional object varChild</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("Office", 9,10,11,12,14,15,16)]
		public void accLocation(out Int32 pxLeft, out Int32 pyTop, out Int32 pcxWidth, out Int32 pcyHeight, object varChild)
		{
			ParameterModifier[] modifiers = Invoker.CreateParamModifiers(true,true,true,true,false);
			pxLeft = 0;
			pyTop = 0;
			pcxWidth = 0;
			pcyHeight = 0;
			object[] paramsArray = Invoker.ValidateParamsArray(pxLeft, pyTop, pcxWidth, pcyHeight, varChild);
			Invoker.Method(this, "accLocation", paramsArray, modifiers);
			pxLeft = (Int32)paramsArray[0];
			pyTop = (Int32)paramsArray[1];
			pcxWidth = (Int32)paramsArray[2];
			pcyHeight = (Int32)paramsArray[3];
		}

		/// <summary>
		/// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="pxLeft">Int32 pxLeft</param>
		/// <param name="pyTop">Int32 pyTop</param>
		/// <param name="pcxWidth">Int32 pcxWidth</param>
		/// <param name="pcyHeight">Int32 pcyHeight</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Office", 9,10,11,12,14,15,16)]
		public void accLocation(out Int32 pxLeft, out Int32 pyTop, out Int32 pcxWidth, out Int32 pcyHeight)
		{
			ParameterModifier[] modifiers = Invoker.CreateParamModifiers(true,true,true,true);
			pxLeft = 0;
			pyTop = 0;
			pcxWidth = 0;
			pcyHeight = 0;
			object[] paramsArray = Invoker.ValidateParamsArray(pxLeft, pyTop, pcxWidth, pcyHeight);
			Invoker.Method(this, "accLocation", paramsArray, modifiers);
			pxLeft = (Int32)paramsArray[0];
			pyTop = (Int32)paramsArray[1];
			pcxWidth = (Int32)paramsArray[2];
			pcyHeight = (Int32)paramsArray[3];
		}

		/// <summary>
		/// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="navDir">Int32 navDir</param>
		/// <param name="varStart">optional object varStart</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("Office", 9,10,11,12,14,15,16)]
		public object accNavigate(Int32 navDir, object varStart)
		{
			return Factory.ExecuteVariantMethodGet(this, "accNavigate", navDir, varStart);
		}

		/// <summary>
		/// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="navDir">Int32 navDir</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Office", 9,10,11,12,14,15,16)]
		public object accNavigate(Int32 navDir)
		{
			return Factory.ExecuteVariantMethodGet(this, "accNavigate", navDir);
		}

		/// <summary>
		/// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="xLeft">Int32 xLeft</param>
		/// <param name="yTop">Int32 yTop</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("Office", 9,10,11,12,14,15,16)]
		public object accHitTest(Int32 xLeft, Int32 yTop)
		{
			return Factory.ExecuteVariantMethodGet(this, "accHitTest", xLeft, yTop);
		}

		/// <summary>
		/// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="varChild">optional object varChild</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("Office", 9,10,11,12,14,15,16)]
		public void accDoDefaultAction(object varChild)
		{
			 Factory.ExecuteMethod(this, "accDoDefaultAction", varChild);
		}

		/// <summary>
		/// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Office", 9,10,11,12,14,15,16)]
		public void accDoDefaultAction()
		{
			 Factory.ExecuteMethod(this, "accDoDefaultAction");
		}

		/// <summary>
		/// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff862231.aspx </remarks>
		[SupportByVersion("Office", 9,10,11,12,14,15,16)]
		public void Delete()
		{
			 Factory.ExecuteMethod(this, "Delete");
		}

		/// <summary>
		/// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff864884.aspx </remarks>
		/// <param name="type">optional object type</param>
		/// <param name="id">optional object id</param>
		/// <param name="tag">optional object tag</param>
		/// <param name="visible">optional object visible</param>
		/// <param name="recursive">optional object recursive</param>
		[SupportByVersion("Office", 9,10,11,12,14,15,16)]
		[BaseResult]
		public NetOffice.OfficeApi.CommandBarControl FindControl(object type, object id, object tag, object visible, object recursive)
		{
			return Factory.ExecuteBaseReferenceMethodGet<NetOffice.OfficeApi.CommandBarControl>(this, "FindControl", new object[]{ type, id, tag, visible, recursive });
		}

		/// <summary>
		/// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff864884.aspx </remarks>
		[CustomMethod]
		[BaseResult]
		[SupportByVersion("Office", 9,10,11,12,14,15,16)]
		public NetOffice.OfficeApi.CommandBarControl FindControl()
		{
			return Factory.ExecuteBaseReferenceMethodGet<NetOffice.OfficeApi.CommandBarControl>(this, "FindControl");
		}

		/// <summary>
		/// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff864884.aspx </remarks>
		/// <param name="type">optional object type</param>
		[CustomMethod]
		[BaseResult]
		[SupportByVersion("Office", 9,10,11,12,14,15,16)]
		public NetOffice.OfficeApi.CommandBarControl FindControl(object type)
		{
			return Factory.ExecuteBaseReferenceMethodGet<NetOffice.OfficeApi.CommandBarControl>(this, "FindControl", type);
		}

		/// <summary>
		/// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff864884.aspx </remarks>
		/// <param name="type">optional object type</param>
		/// <param name="id">optional object id</param>
		[CustomMethod]
		[BaseResult]
		[SupportByVersion("Office", 9,10,11,12,14,15,16)]
		public NetOffice.OfficeApi.CommandBarControl FindControl(object type, object id)
		{
			return Factory.ExecuteBaseReferenceMethodGet<NetOffice.OfficeApi.CommandBarControl>(this, "FindControl", type, id);
		}

		/// <summary>
		/// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff864884.aspx </remarks>
		/// <param name="type">optional object type</param>
		/// <param name="id">optional object id</param>
		/// <param name="tag">optional object tag</param>
		[CustomMethod]
		[BaseResult]
		[SupportByVersion("Office", 9,10,11,12,14,15,16)]
		public NetOffice.OfficeApi.CommandBarControl FindControl(object type, object id, object tag)
		{
			return Factory.ExecuteBaseReferenceMethodGet<NetOffice.OfficeApi.CommandBarControl>(this, "FindControl", type, id, tag);
		}

		/// <summary>
		/// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff864884.aspx </remarks>
		/// <param name="type">optional object type</param>
		/// <param name="id">optional object id</param>
		/// <param name="tag">optional object tag</param>
		/// <param name="visible">optional object visible</param>
		[CustomMethod]
		[BaseResult]
		[SupportByVersion("Office", 9,10,11,12,14,15,16)]
		public NetOffice.OfficeApi.CommandBarControl FindControl(object type, object id, object tag, object visible)
		{
			return Factory.ExecuteBaseReferenceMethodGet<NetOffice.OfficeApi.CommandBarControl>(this, "FindControl", type, id, tag, visible);
		}

		/// <summary>
		/// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff863143.aspx </remarks>
		[SupportByVersion("Office", 9,10,11,12,14,15,16)]
		public void Reset()
		{
			 Factory.ExecuteMethod(this, "Reset");
		}

		/// <summary>
		/// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff865188.aspx </remarks>
		/// <param name="x">optional object x</param>
		/// <param name="y">optional object y</param>
		[SupportByVersion("Office", 9,10,11,12,14,15,16)]
		public void ShowPopup(object x, object y)
		{
			 Factory.ExecuteMethod(this, "ShowPopup", x, y);
		}

		/// <summary>
		/// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff865188.aspx </remarks>
		[CustomMethod]
		[SupportByVersion("Office", 9,10,11,12,14,15,16)]
		public void ShowPopup()
		{
			 Factory.ExecuteMethod(this, "ShowPopup");
		}

		/// <summary>
		/// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff865188.aspx </remarks>
		/// <param name="x">optional object x</param>
		[CustomMethod]
		[SupportByVersion("Office", 9,10,11,12,14,15,16)]
		public void ShowPopup(object x)
		{
			 Factory.ExecuteMethod(this, "ShowPopup", x);
		}

		#endregion

		#pragma warning restore
	}
}
