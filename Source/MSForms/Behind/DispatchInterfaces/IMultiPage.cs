using System.Reflection;
using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;
using NetOffice.MSFormsApi;

namespace NetOffice.MSFormsApi.Behind
{
	/// <summary>
	/// DispatchInterface IMultiPage 
	/// SupportByVersion MSForms, 2
	/// </summary>
	[SupportByVersion("MSForms", 2)]
	[EntityType(EntityType.IsDispatchInterface), BaseType]
 	public class IMultiPage : COMObject, NetOffice.MSFormsApi.IMultiPage
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
                    _contractType = typeof(NetOffice.MSFormsApi.IMultiPage);
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
                    _type = typeof(IMultiPage);
                return _type;
            }
        }
        
        #endregion
        
		#region Ctor

		/// <summary>
		/// Stub Ctor, not indented to use
		/// </summary>
		public IMultiPage() : base()
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
		/// Get
		/// </summary>
		[SupportByVersion("MSForms", 2)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public virtual NetOffice.MSFormsApi.Page SelectedItem
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.MSFormsApi.Page>(this, "SelectedItem", typeof(NetOffice.MSFormsApi.Page));
			}
		}

		/// <summary>
		/// SupportByVersion MSForms 2
		/// Get
		/// </summary>
		[SupportByVersion("MSForms", 2)]
		public virtual NetOffice.MSFormsApi.Pages Pages
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.MSFormsApi.Pages>(this, "Pages", typeof(NetOffice.MSFormsApi.Pages));
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
		/// <param name="width">Int32 width</param>
		[SupportByVersion("MSForms", 2)]
		public virtual void _SetTabFixedWidth(Int32 width)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "_SetTabFixedWidth", width);
		}

		/// <summary>
		/// SupportByVersion MSForms 2
		/// </summary>
		/// <param name="width">Int32 width</param>
		[SupportByVersion("MSForms", 2)]
		public virtual void _GetTabFixedWidth(out Int32 width)
		{
			ParameterModifier[] modifiers = Invoker.CreateParamModifiers(true);
			width = 0;
			object[] paramsArray = Invoker.ValidateParamsArray(width);
			Invoker.Method(this, "_GetTabFixedWidth", paramsArray, modifiers);
			width = (Int32)paramsArray[0];
		}

		/// <summary>
		/// SupportByVersion MSForms 2
		/// </summary>
		/// <param name="height">Int32 height</param>
		[SupportByVersion("MSForms", 2)]
		public virtual void _SetTabFixedHeight(Int32 height)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "_SetTabFixedHeight", height);
		}

		/// <summary>
		/// SupportByVersion MSForms 2
		/// </summary>
		/// <param name="height">Int32 height</param>
		[SupportByVersion("MSForms", 2)]
		public virtual void _GetTabFixedHeight(out Int32 height)
		{
			ParameterModifier[] modifiers = Invoker.CreateParamModifiers(true);
			height = 0;
			object[] paramsArray = Invoker.ValidateParamsArray(height);
			Invoker.Method(this, "_GetTabFixedHeight", paramsArray, modifiers);
			height = (Int32)paramsArray[0];
		}

		#endregion

		#pragma warning restore
	}
}
