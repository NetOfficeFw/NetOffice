using System.Reflection;
using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;
using NetOffice.MSFormsApi;

namespace NetOffice.MSFormsApi.Behind
{
	/// <summary>
	/// DispatchInterface ITabStrip 
	/// SupportByVersion MSForms, 2
	/// </summary>
	[SupportByVersion("MSForms", 2)]
	[EntityType(EntityType.IsDispatchInterface), BaseType]
 	public class ITabStrip : COMObject, NetOffice.MSFormsApi.ITabStrip
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
                    _contractType = typeof(NetOffice.MSFormsApi.ITabStrip);
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
                    _type = typeof(ITabStrip);
                return _type;
            }
        }
        
        #endregion
        
		#region Ctor

		/// <summary>
		/// Stub Ctor, not indented to use
		/// </summary>
		public ITabStrip() : base()
		{

		}

		#endregion
		
		#region Properties

		/// <summary>
		/// SupportByVersion MSForms 2
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSForms", 2)]
		public virtual Int32 BackColor
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteInt32PropertyGet(this, "BackColor");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "BackColor", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSForms 2
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSForms", 2)]
		public virtual Int32 ForeColor
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteInt32PropertyGet(this, "ForeColor");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "ForeColor", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSForms 2
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSForms", 2)]
		[BaseResult]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public virtual NetOffice.MSFormsApi.Font _Font_Reserved
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteBaseReferencePropertyGet<NetOffice.MSFormsApi.Font>(this, "_Font_Reserved");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteReferencePropertySet(this, "_Font_Reserved", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSForms 2
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSForms", 2)]
		[BaseResult]
		public virtual NetOffice.MSFormsApi.Font Font
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteBaseReferencePropertyGet<NetOffice.MSFormsApi.Font>(this, "Font");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteReferencePropertySet(this, "Font", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSForms 2
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSForms", 2)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public virtual string FontName
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "FontName");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "FontName", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSForms 2
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSForms", 2)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public virtual bool FontBold
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "FontBold");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "FontBold", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSForms 2
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSForms", 2)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public virtual bool FontItalic
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "FontItalic");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "FontItalic", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSForms 2
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSForms", 2)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public virtual bool FontUnderline
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "FontUnderline");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "FontUnderline", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSForms 2
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSForms", 2)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public virtual bool FontStrikethru
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "FontStrikethru");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "FontStrikethru", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSForms 2
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSForms", 2)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public virtual float FontSize
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteFloatPropertyGet(this, "FontSize");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "FontSize", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSForms 2
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSForms", 2)]
		public virtual bool Enabled
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "Enabled");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "Enabled", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSForms 2
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSForms", 2), NativeResult]
		public virtual stdole.Picture MouseIcon
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
		public virtual NetOffice.MSFormsApi.Enums.fmMousePointer MousePointer
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteEnumPropertyGet<NetOffice.MSFormsApi.Enums.fmMousePointer>(this, "MousePointer");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteEnumPropertySet(this, "MousePointer", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSForms 2
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSForms", 2)]
		public virtual bool MultiRow
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "MultiRow");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "MultiRow", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSForms 2
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSForms", 2)]
		public virtual NetOffice.MSFormsApi.Enums.fmTabStyle Style
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteEnumPropertyGet<NetOffice.MSFormsApi.Enums.fmTabStyle>(this, "Style");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteEnumPropertySet(this, "Style", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSForms 2
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSForms", 2)]
		public virtual NetOffice.MSFormsApi.Enums.fmTabOrientation TabOrientation
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteEnumPropertyGet<NetOffice.MSFormsApi.Enums.fmTabOrientation>(this, "TabOrientation");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteEnumPropertySet(this, "TabOrientation", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSForms 2
		/// Get
		/// </summary>
		[SupportByVersion("MSForms", 2)]
		public virtual Single ClientTop
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteSinglePropertyGet(this, "ClientTop");
			}
		}

		/// <summary>
		/// SupportByVersion MSForms 2
		/// Get
		/// </summary>
		[SupportByVersion("MSForms", 2)]
		public virtual Single ClientLeft
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteSinglePropertyGet(this, "ClientLeft");
			}
		}

		/// <summary>
		/// SupportByVersion MSForms 2
		/// Get
		/// </summary>
		[SupportByVersion("MSForms", 2)]
		public virtual Single ClientWidth
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteSinglePropertyGet(this, "ClientWidth");
			}
		}

		/// <summary>
		/// SupportByVersion MSForms 2
		/// Get
		/// </summary>
		[SupportByVersion("MSForms", 2)]
		public virtual Single ClientHeight
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteSinglePropertyGet(this, "ClientHeight");
			}
		}

		/// <summary>
		/// SupportByVersion MSForms 2
		/// Get
		/// </summary>
		[SupportByVersion("MSForms", 2)]
		public virtual NetOffice.MSFormsApi.Tabs Tabs
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.MSFormsApi.Tabs>(this, "Tabs", typeof(NetOffice.MSFormsApi.Tabs));
			}
		}

		/// <summary>
		/// SupportByVersion MSForms 2
		/// Get
		/// </summary>
		[SupportByVersion("MSForms", 2)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public virtual NetOffice.MSFormsApi.Tab SelectedItem
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.MSFormsApi.Tab>(this, "SelectedItem", typeof(NetOffice.MSFormsApi.Tab));
			}
		}

		/// <summary>
		/// SupportByVersion MSForms 2
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSForms", 2)]
		public virtual Int32 Value
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteInt32PropertyGet(this, "Value");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "Value", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSForms 2
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSForms", 2)]
		public virtual Single TabFixedWidth
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteSinglePropertyGet(this, "TabFixedWidth");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "TabFixedWidth", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSForms 2
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSForms", 2)]
		public virtual Single TabFixedHeight
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteSinglePropertyGet(this, "TabFixedHeight");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "TabFixedHeight", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSForms 2
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSForms", 2)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public virtual Int16 FontWeight
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteInt16PropertyGet(this, "FontWeight");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "FontWeight", value);
			}
		}

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion MSForms 2
		/// </summary>
		/// <param name="tabFixedWidth">Int32 tabFixedWidth</param>
		[SupportByVersion("MSForms", 2)]
		public virtual void _SetTabFixedWidth(Int32 tabFixedWidth)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "_SetTabFixedWidth", tabFixedWidth);
		}

		/// <summary>
		/// SupportByVersion MSForms 2
		/// </summary>
		/// <param name="tabFixedWidth">Int32 tabFixedWidth</param>
		[SupportByVersion("MSForms", 2)]
		public virtual void _GetTabFixedWidth(out Int32 tabFixedWidth)
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
		public virtual void _SetTabFixedHeight(Int32 tabFixedHeight)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "_SetTabFixedHeight", tabFixedHeight);
		}

		/// <summary>
		/// SupportByVersion MSForms 2
		/// </summary>
		/// <param name="tabFixedHeight">Int32 tabFixedHeight</param>
		[SupportByVersion("MSForms", 2)]
		public virtual void _GetTabFixedHeight(out Int32 tabFixedHeight)
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
		public virtual void _GetClientTop(out Int32 clientTop)
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
		public virtual void _GetClientLeft(out Int32 clientLeft)
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
		public virtual void _GetClientWidth(out Int32 clientWidth)
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
		public virtual void _GetClientHeight(out Int32 clientHeight)
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


