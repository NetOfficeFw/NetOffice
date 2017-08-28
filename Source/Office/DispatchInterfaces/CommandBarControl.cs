using System.Reflection;
using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.OfficeApi
{
	/// <summary>
	/// CommandBarControl
	/// </summary>
	[SyntaxBypass]
 	public class CommandBarControl_ : _IMsoOleAccDispObj
	{
		#region Construction

		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
		public CommandBarControl_(Core factory, ICOMObject parentObject, object comProxy) : base(factory, parentObject, comProxy)
		{
			
		}

        /// <param name="parentObject">object there has created the proxy</param>
        /// <param name="comProxy">inner wrapped COM proxy</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public CommandBarControl_(ICOMObject parentObject, object comProxy) : base(parentObject, comProxy)
		{
		}
		
		/// <param name="parentObject">object there has created the proxy</param>
        /// <param name="comProxy">inner wrapped COM proxy</param>
        /// <param name="comProxyType">Type of inner wrapped COM proxy"</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public CommandBarControl_(ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(parentObject, comProxy, comProxyType)
		{
		}

		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public CommandBarControl_(Core factory, ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(factory, parentObject, comProxy, comProxyType)
		{

		}
		
		/// <param name="replacedObject">object to replaced. replacedObject are not usable after this action</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public CommandBarControl_(ICOMObject replacedObject) : base(replacedObject)
		{
		}

		/// <summary>
        /// Hidden stub .ctor
        /// </summary>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public CommandBarControl_() : base()
		{
		}
		
		/// <param name="progId">registered progID</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public CommandBarControl_(string progId) : base(progId)
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
	/// DispatchInterface CommandBarControl 
	/// SupportByVersion Office, 9,10,11,12,14,15,16
	/// </summary>
	/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff863821.aspx </remarks>
	[SupportByVersion("Office", 9,10,11,12,14,15,16)]
	[EntityType(EntityType.IsDispatchInterface)]
 	public class CommandBarControl : CommandBarControl_
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
                    _type = typeof(CommandBarControl);
                return _type;
            }
        }
        
        #endregion
        
		#region Construction

		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
		public CommandBarControl(Core factory, ICOMObject parentObject, object comProxy) : base(factory, parentObject, comProxy)
		{
			
		}

        ///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public CommandBarControl(ICOMObject parentObject, object comProxy) : base(parentObject, comProxy)
		{
		}
		
		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public CommandBarControl(Core factory, ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(factory, parentObject, comProxy, comProxyType)
		{

		}

		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public CommandBarControl(ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(parentObject, comProxy, comProxyType)
		{
		}
		
		///<param name="replacedObject">object to replaced. replacedObject are not usable after this action</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public CommandBarControl(ICOMObject replacedObject) : base(replacedObject)
		{
		}
		
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public CommandBarControl() : base()
		{
		}
		
		/// <param name="progId">registered progID</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public CommandBarControl(string progId) : base(progId)
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
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff861759.aspx </remarks>
		[SupportByVersion("Office", 9,10,11,12,14,15,16)]
		public bool BeginGroup
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "BeginGroup");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "BeginGroup", value);
			}
		}

		/// <summary>
		/// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff861515.aspx </remarks>
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
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff862411.aspx </remarks>
		[SupportByVersion("Office", 9,10,11,12,14,15,16)]
		public string Caption
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "Caption");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "Caption", value);
			}
		}

		/// <summary>
		/// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		[SupportByVersion("Office", 9,10,11,12,14,15,16), ProxyResult]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public object Control
		{
			get
			{
				return Factory.ExecuteReferencePropertyGet(this, "Control");
			}
		}

		/// <summary>
		/// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff861718.aspx </remarks>
		[SupportByVersion("Office", 9,10,11,12,14,15,16)]
		public string DescriptionText
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "DescriptionText");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "DescriptionText", value);
			}
		}

		/// <summary>
		/// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff862488.aspx </remarks>
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
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff862462.aspx </remarks>
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
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff861795.aspx </remarks>
		[SupportByVersion("Office", 9,10,11,12,14,15,16)]
		public Int32 HelpContextId
		{
			get
			{
				return Factory.ExecuteInt32PropertyGet(this, "HelpContextId");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "HelpContextId", value);
			}
		}

		/// <summary>
		/// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff860796.aspx </remarks>
		[SupportByVersion("Office", 9,10,11,12,14,15,16)]
		public string HelpFile
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "HelpFile");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "HelpFile", value);
			}
		}

		/// <summary>
		/// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff860277.aspx </remarks>
		[SupportByVersion("Office", 9,10,11,12,14,15,16)]
		public Int32 Id
		{
			get
			{
				return Factory.ExecuteInt32PropertyGet(this, "Id");
			}
		}

		/// <summary>
		/// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff860502.aspx </remarks>
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
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff861878.aspx </remarks>
		[SupportByVersion("Office", 9,10,11,12,14,15,16)]
		public Int32 Left
		{
			get
			{
				return Factory.ExecuteInt32PropertyGet(this, "Left");
			}
		}

		/// <summary>
		/// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff864177.aspx </remarks>
		[SupportByVersion("Office", 9,10,11,12,14,15,16)]
		public NetOffice.OfficeApi.Enums.MsoControlOLEUsage OLEUsage
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.OfficeApi.Enums.MsoControlOLEUsage>(this, "OLEUsage");
			}
			set
			{
				Factory.ExecuteEnumPropertySet(this, "OLEUsage", value);
			}
		}

		/// <summary>
		/// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff860249.aspx </remarks>
		[SupportByVersion("Office", 9,10,11,12,14,15,16)]
		public string OnAction
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "OnAction");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "OnAction", value);
			}
		}

		/// <summary>
		/// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff864892.aspx </remarks>
		[SupportByVersion("Office", 9,10,11,12,14,15,16)]
		public NetOffice.OfficeApi.CommandBar Parent
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.OfficeApi.CommandBar>(this, "Parent", NetOffice.OfficeApi.CommandBar.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff862239.aspx </remarks>
		[SupportByVersion("Office", 9,10,11,12,14,15,16)]
		public string Parameter
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "Parameter");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "Parameter", value);
			}
		}

		/// <summary>
		/// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff860592.aspx </remarks>
		[SupportByVersion("Office", 9,10,11,12,14,15,16)]
		public Int32 Priority
		{
			get
			{
				return Factory.ExecuteInt32PropertyGet(this, "Priority");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "Priority", value);
			}
		}

		/// <summary>
		/// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff864872.aspx </remarks>
		[SupportByVersion("Office", 9,10,11,12,14,15,16)]
		public string Tag
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "Tag");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "Tag", value);
			}
		}

		/// <summary>
		/// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff860239.aspx </remarks>
		[SupportByVersion("Office", 9,10,11,12,14,15,16)]
		public string TooltipText
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "TooltipText");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "TooltipText", value);
			}
		}

		/// <summary>
		/// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff862470.aspx </remarks>
		[SupportByVersion("Office", 9,10,11,12,14,15,16)]
		public Int32 Top
		{
			get
			{
				return Factory.ExecuteInt32PropertyGet(this, "Top");
			}
		}

		/// <summary>
		/// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff863407.aspx </remarks>
		[SupportByVersion("Office", 9,10,11,12,14,15,16)]
		public NetOffice.OfficeApi.Enums.MsoControlType Type
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.OfficeApi.Enums.MsoControlType>(this, "Type");
			}
		}

		/// <summary>
		/// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff863300.aspx </remarks>
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
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff863521.aspx </remarks>
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
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff864638.aspx </remarks>
		[SupportByVersion("Office", 9,10,11,12,14,15,16)]
		public bool IsPriorityDropped
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "IsPriorityDropped");
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
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff861397.aspx </remarks>
		/// <param name="bar">optional object bar</param>
		/// <param name="before">optional object before</param>
		[SupportByVersion("Office", 9,10,11,12,14,15,16)]
		public NetOffice.OfficeApi.CommandBarControl Copy(object bar, object before)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.OfficeApi.CommandBarControl>(this, "Copy", NetOffice.OfficeApi.CommandBarControl.LateBindingApiWrapperType, bar, before);
		}

		/// <summary>
		/// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff861397.aspx </remarks>
		[CustomMethod]
		[SupportByVersion("Office", 9,10,11,12,14,15,16)]
		public NetOffice.OfficeApi.CommandBarControl Copy()
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.OfficeApi.CommandBarControl>(this, "Copy", NetOffice.OfficeApi.CommandBarControl.LateBindingApiWrapperType);
		}

		/// <summary>
		/// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff861397.aspx </remarks>
		/// <param name="bar">optional object bar</param>
		[CustomMethod]
		[SupportByVersion("Office", 9,10,11,12,14,15,16)]
		public NetOffice.OfficeApi.CommandBarControl Copy(object bar)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.OfficeApi.CommandBarControl>(this, "Copy", NetOffice.OfficeApi.CommandBarControl.LateBindingApiWrapperType, bar);
		}

		/// <summary>
		/// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff865269.aspx </remarks>
		/// <param name="temporary">optional object temporary</param>
		[SupportByVersion("Office", 9,10,11,12,14,15,16)]
		public void Delete(object temporary)
		{
			 Factory.ExecuteMethod(this, "Delete", temporary);
		}

		/// <summary>
		/// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff865269.aspx </remarks>
		[CustomMethod]
		[SupportByVersion("Office", 9,10,11,12,14,15,16)]
		public void Delete()
		{
			 Factory.ExecuteMethod(this, "Delete");
		}

		/// <summary>
		/// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff861885.aspx </remarks>
		[SupportByVersion("Office", 9,10,11,12,14,15,16)]
		public void Execute()
		{
			 Factory.ExecuteMethod(this, "Execute");
		}

		/// <summary>
		/// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff863114.aspx </remarks>
		/// <param name="bar">optional object bar</param>
		/// <param name="before">optional object before</param>
		[SupportByVersion("Office", 9,10,11,12,14,15,16)]
		public NetOffice.OfficeApi.CommandBarControl Move(object bar, object before)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.OfficeApi.CommandBarControl>(this, "Move", NetOffice.OfficeApi.CommandBarControl.LateBindingApiWrapperType, bar, before);
		}

		/// <summary>
		/// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff863114.aspx </remarks>
		[CustomMethod]
		[SupportByVersion("Office", 9,10,11,12,14,15,16)]
		public NetOffice.OfficeApi.CommandBarControl Move()
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.OfficeApi.CommandBarControl>(this, "Move", NetOffice.OfficeApi.CommandBarControl.LateBindingApiWrapperType);
		}

		/// <summary>
		/// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff863114.aspx </remarks>
		/// <param name="bar">optional object bar</param>
		[CustomMethod]
		[SupportByVersion("Office", 9,10,11,12,14,15,16)]
		public NetOffice.OfficeApi.CommandBarControl Move(object bar)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.OfficeApi.CommandBarControl>(this, "Move", NetOffice.OfficeApi.CommandBarControl.LateBindingApiWrapperType, bar);
		}

		/// <summary>
		/// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff862736.aspx </remarks>
		[SupportByVersion("Office", 9,10,11,12,14,15,16)]
		public void Reset()
		{
			 Factory.ExecuteMethod(this, "Reset");
		}

		/// <summary>
		/// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff864985.aspx </remarks>
		[SupportByVersion("Office", 9,10,11,12,14,15,16)]
		public void SetFocus()
		{
			 Factory.ExecuteMethod(this, "SetFocus");
		}

		/// <summary>
		/// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("Office", 9,10,11,12,14,15,16)]
		public void Reserved1()
		{
			 Factory.ExecuteMethod(this, "Reserved1");
		}

		/// <summary>
		/// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("Office", 9,10,11,12,14,15,16)]
		public void Reserved2()
		{
			 Factory.ExecuteMethod(this, "Reserved2");
		}

		/// <summary>
		/// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("Office", 9,10,11,12,14,15,16)]
		public void Reserved3()
		{
			 Factory.ExecuteMethod(this, "Reserved3");
		}

		/// <summary>
		/// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("Office", 9,10,11,12,14,15,16)]
		public void Reserved4()
		{
			 Factory.ExecuteMethod(this, "Reserved4");
		}

		/// <summary>
		/// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("Office", 9,10,11,12,14,15,16)]
		public void Reserved5()
		{
			 Factory.ExecuteMethod(this, "Reserved5");
		}

		/// <summary>
		/// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("Office", 9,10,11,12,14,15,16)]
		public void Reserved6()
		{
			 Factory.ExecuteMethod(this, "Reserved6");
		}

		/// <summary>
		/// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("Office", 9,10,11,12,14,15,16)]
		public void Reserved7()
		{
			 Factory.ExecuteMethod(this, "Reserved7");
		}

		#endregion

		#pragma warning restore
	}
}



