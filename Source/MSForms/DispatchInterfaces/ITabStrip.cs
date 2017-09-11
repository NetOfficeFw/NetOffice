using System.Reflection;
using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.MSFormsApi
{
	/// <summary>
	/// DispatchInterface ITabStrip 
	/// SupportByVersion MSForms, 2
	/// </summary>
	[SupportByVersion("MSForms", 2)]
	[EntityType(EntityType.IsDispatchInterface), BaseType]
 	public class ITabStrip : COMObject
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
                    _type = typeof(ITabStrip);
                return _type;
            }
        }
        
        #endregion
        
		#region Ctor

		/// <param name="factory">current used factory core</param>
		/// <param name="parentObject">object there has created the proxy</param>
		/// <param name="proxyShare">proxy share instead if com proxy</param>
		public ITabStrip(Core factory, ICOMObject parentObject, COMProxyShare proxyShare) : base(factory, parentObject, proxyShare)
		{
		}

		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
		public ITabStrip(Core factory, ICOMObject parentObject, object comProxy) : base(factory, parentObject, comProxy)
		{
			
		}

        ///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public ITabStrip(ICOMObject parentObject, object comProxy) : base(parentObject, comProxy)
		{
		}
		
		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public ITabStrip(Core factory, ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(factory, parentObject, comProxy, comProxyType)
		{

		}

		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public ITabStrip(ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(parentObject, comProxy, comProxyType)
		{
		}
		
		///<param name="replacedObject">object to replaced. replacedObject are not usable after this action</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public ITabStrip(ICOMObject replacedObject) : base(replacedObject)
		{
		}
		
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public ITabStrip() : base()
		{
		}
		
		/// <param name="progId">registered progID</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public ITabStrip(string progId) : base(progId)
		{
		}
		
		#endregion
		
		#region Properties

		/// <summary>
		/// SupportByVersion MSForms 2
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSForms", 2)]
		public Int32 BackColor
		{
			get
			{
				return Factory.ExecuteInt32PropertyGet(this, "BackColor");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "BackColor", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSForms 2
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSForms", 2)]
		public Int32 ForeColor
		{
			get
			{
				return Factory.ExecuteInt32PropertyGet(this, "ForeColor");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "ForeColor", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSForms 2
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSForms", 2)]
		[BaseResult]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public NetOffice.MSFormsApi.Font _Font_Reserved
		{
			get
			{
				return Factory.ExecuteBaseReferencePropertyGet<NetOffice.MSFormsApi.Font>(this, "_Font_Reserved");
			}
			set
			{
				Factory.ExecuteReferencePropertySet(this, "_Font_Reserved", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSForms 2
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSForms", 2)]
		[BaseResult]
		public NetOffice.MSFormsApi.Font Font
		{
			get
			{
				return Factory.ExecuteBaseReferencePropertyGet<NetOffice.MSFormsApi.Font>(this, "Font");
			}
			set
			{
				Factory.ExecuteReferencePropertySet(this, "Font", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSForms 2
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSForms", 2)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public string FontName
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "FontName");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "FontName", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSForms 2
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSForms", 2)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public bool FontBold
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "FontBold");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "FontBold", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSForms 2
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSForms", 2)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public bool FontItalic
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "FontItalic");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "FontItalic", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSForms 2
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSForms", 2)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public bool FontUnderline
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "FontUnderline");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "FontUnderline", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSForms 2
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSForms", 2)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public bool FontStrikethru
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "FontStrikethru");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "FontStrikethru", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSForms 2
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSForms", 2)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public float FontSize
		{
			get
			{
				return Factory.ExecuteFloatPropertyGet(this, "FontSize");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "FontSize", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSForms 2
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSForms", 2)]
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
		/// SupportByVersion MSForms 2
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSForms", 2), NativeResult]
		public stdole.Picture MouseIcon
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "MouseIcon", paramsArray);
                return returnItem as stdole.Picture;
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "MouseIcon", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion MSForms 2
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSForms", 2)]
		public NetOffice.MSFormsApi.Enums.fmMousePointer MousePointer
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.MSFormsApi.Enums.fmMousePointer>(this, "MousePointer");
			}
			set
			{
				Factory.ExecuteEnumPropertySet(this, "MousePointer", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSForms 2
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSForms", 2)]
		public bool MultiRow
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "MultiRow");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "MultiRow", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSForms 2
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSForms", 2)]
		public NetOffice.MSFormsApi.Enums.fmTabStyle Style
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.MSFormsApi.Enums.fmTabStyle>(this, "Style");
			}
			set
			{
				Factory.ExecuteEnumPropertySet(this, "Style", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSForms 2
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSForms", 2)]
		public NetOffice.MSFormsApi.Enums.fmTabOrientation TabOrientation
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.MSFormsApi.Enums.fmTabOrientation>(this, "TabOrientation");
			}
			set
			{
				Factory.ExecuteEnumPropertySet(this, "TabOrientation", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSForms 2
		/// Get
		/// </summary>
		[SupportByVersion("MSForms", 2)]
		public Single ClientTop
		{
			get
			{
				return Factory.ExecuteSinglePropertyGet(this, "ClientTop");
			}
		}

		/// <summary>
		/// SupportByVersion MSForms 2
		/// Get
		/// </summary>
		[SupportByVersion("MSForms", 2)]
		public Single ClientLeft
		{
			get
			{
				return Factory.ExecuteSinglePropertyGet(this, "ClientLeft");
			}
		}

		/// <summary>
		/// SupportByVersion MSForms 2
		/// Get
		/// </summary>
		[SupportByVersion("MSForms", 2)]
		public Single ClientWidth
		{
			get
			{
				return Factory.ExecuteSinglePropertyGet(this, "ClientWidth");
			}
		}

		/// <summary>
		/// SupportByVersion MSForms 2
		/// Get
		/// </summary>
		[SupportByVersion("MSForms", 2)]
		public Single ClientHeight
		{
			get
			{
				return Factory.ExecuteSinglePropertyGet(this, "ClientHeight");
			}
		}

		/// <summary>
		/// SupportByVersion MSForms 2
		/// Get
		/// </summary>
		[SupportByVersion("MSForms", 2)]
		public NetOffice.MSFormsApi.Tabs Tabs
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.MSFormsApi.Tabs>(this, "Tabs", NetOffice.MSFormsApi.Tabs.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion MSForms 2
		/// Get
		/// </summary>
		[SupportByVersion("MSForms", 2)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public NetOffice.MSFormsApi.Tab SelectedItem
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.MSFormsApi.Tab>(this, "SelectedItem", NetOffice.MSFormsApi.Tab.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion MSForms 2
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSForms", 2)]
		public Int32 Value
		{
			get
			{
				return Factory.ExecuteInt32PropertyGet(this, "Value");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "Value", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSForms 2
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSForms", 2)]
		public Single TabFixedWidth
		{
			get
			{
				return Factory.ExecuteSinglePropertyGet(this, "TabFixedWidth");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "TabFixedWidth", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSForms 2
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSForms", 2)]
		public Single TabFixedHeight
		{
			get
			{
				return Factory.ExecuteSinglePropertyGet(this, "TabFixedHeight");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "TabFixedHeight", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSForms 2
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSForms", 2)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Int16 FontWeight
		{
			get
			{
				return Factory.ExecuteInt16PropertyGet(this, "FontWeight");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "FontWeight", value);
			}
		}

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion MSForms 2
		/// </summary>
		/// <param name="tabFixedWidth">Int32 tabFixedWidth</param>
		[SupportByVersion("MSForms", 2)]
		public void _SetTabFixedWidth(Int32 tabFixedWidth)
		{
			 Factory.ExecuteMethod(this, "_SetTabFixedWidth", tabFixedWidth);
		}

		/// <summary>
		/// SupportByVersion MSForms 2
		/// </summary>
		/// <param name="tabFixedWidth">Int32 tabFixedWidth</param>
		[SupportByVersion("MSForms", 2)]
		public void _GetTabFixedWidth(out Int32 tabFixedWidth)
		{
			ParameterModifier[] modifiers = Invoker.CreateParamModifiers(true);
			tabFixedWidth = 0;
			object[] paramsArray = Invoker.ValidateParamsArray(tabFixedWidth);
			Invoker.Method(this, "_GetTabFixedWidth", paramsArray, modifiers);
			tabFixedWidth = (Int32)paramsArray[0];
		}

		/// <summary>
		/// SupportByVersion MSForms 2
		/// </summary>
		/// <param name="tabFixedHeight">Int32 tabFixedHeight</param>
		[SupportByVersion("MSForms", 2)]
		public void _SetTabFixedHeight(Int32 tabFixedHeight)
		{
			 Factory.ExecuteMethod(this, "_SetTabFixedHeight", tabFixedHeight);
		}

		/// <summary>
		/// SupportByVersion MSForms 2
		/// </summary>
		/// <param name="tabFixedHeight">Int32 tabFixedHeight</param>
		[SupportByVersion("MSForms", 2)]
		public void _GetTabFixedHeight(out Int32 tabFixedHeight)
		{
			ParameterModifier[] modifiers = Invoker.CreateParamModifiers(true);
			tabFixedHeight = 0;
			object[] paramsArray = Invoker.ValidateParamsArray(tabFixedHeight);
			Invoker.Method(this, "_GetTabFixedHeight", paramsArray, modifiers);
			tabFixedHeight = (Int32)paramsArray[0];
		}

		/// <summary>
		/// SupportByVersion MSForms 2
		/// </summary>
		/// <param name="clientTop">Int32 clientTop</param>
		[SupportByVersion("MSForms", 2)]
		public void _GetClientTop(out Int32 clientTop)
		{
			ParameterModifier[] modifiers = Invoker.CreateParamModifiers(true);
			clientTop = 0;
			object[] paramsArray = Invoker.ValidateParamsArray(clientTop);
			Invoker.Method(this, "_GetClientTop", paramsArray, modifiers);
			clientTop = (Int32)paramsArray[0];
		}

		/// <summary>
		/// SupportByVersion MSForms 2
		/// </summary>
		/// <param name="clientLeft">Int32 clientLeft</param>
		[SupportByVersion("MSForms", 2)]
		public void _GetClientLeft(out Int32 clientLeft)
		{
			ParameterModifier[] modifiers = Invoker.CreateParamModifiers(true);
			clientLeft = 0;
			object[] paramsArray = Invoker.ValidateParamsArray(clientLeft);
			Invoker.Method(this, "_GetClientLeft", paramsArray, modifiers);
			clientLeft = (Int32)paramsArray[0];
		}

		/// <summary>
		/// SupportByVersion MSForms 2
		/// </summary>
		/// <param name="clientWidth">Int32 clientWidth</param>
		[SupportByVersion("MSForms", 2)]
		public void _GetClientWidth(out Int32 clientWidth)
		{
			ParameterModifier[] modifiers = Invoker.CreateParamModifiers(true);
			clientWidth = 0;
			object[] paramsArray = Invoker.ValidateParamsArray(clientWidth);
			Invoker.Method(this, "_GetClientWidth", paramsArray, modifiers);
			clientWidth = (Int32)paramsArray[0];
		}

		/// <summary>
		/// SupportByVersion MSForms 2
		/// </summary>
		/// <param name="clientHeight">Int32 clientHeight</param>
		[SupportByVersion("MSForms", 2)]
		public void _GetClientHeight(out Int32 clientHeight)
		{
			ParameterModifier[] modifiers = Invoker.CreateParamModifiers(true);
			clientHeight = 0;
			object[] paramsArray = Invoker.ValidateParamsArray(clientHeight);
			Invoker.Method(this, "_GetClientHeight", paramsArray, modifiers);
			clientHeight = (Int32)paramsArray[0];
		}

		#endregion

		#pragma warning restore
	}
}
