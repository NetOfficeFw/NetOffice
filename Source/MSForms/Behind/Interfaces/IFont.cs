using System.Reflection;
using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;
using NetOffice.MSFormsApi;

namespace NetOffice.MSFormsApi.Behind
{
	/// <summary>
	/// Interface IFont 
	/// SupportByVersion MSForms, 2
	/// </summary>
	[SupportByVersion("MSForms", 2)]
	[EntityType(EntityType.IsInterface)]
 	public class IFont : COMObject, NetOffice.MSFormsApi.IFont
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
                    _contractType = typeof(NetOffice.MSFormsApi.IFont);
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
                    _type = typeof(IFont);
                return _type;
            }
        }
        
        #endregion
        
		#region Ctor

		/// <summary>
		/// Stub Ctor, not indented to use
		/// </summary>
		public IFont() : base()
		{

		}

		#endregion
		
		#region Properties

		/// <summary>
		/// SupportByVersion MSForms 2
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSForms", 2)]
		public virtual string Name
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "Name");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "Name", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSForms 2
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSForms", 2)]
		public virtual float Size
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteFloatPropertyGet(this, "Size");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "Size", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSForms 2
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSForms", 2)]
		public virtual bool Bold
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "Bold");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "Bold", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSForms 2
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSForms", 2)]
		public virtual bool Italic
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "Italic");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "Italic", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSForms 2
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSForms", 2)]
		public virtual bool Underline
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "Underline");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "Underline", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSForms 2
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSForms", 2)]
		public virtual bool Strikethrough
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "Strikethrough");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "Strikethrough", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSForms 2
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSForms", 2)]
		public virtual Int16 Weight
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteInt16PropertyGet(this, "Weight");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "Weight", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSForms 2
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSForms", 2)]
		public virtual Int16 Charset
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteInt16PropertyGet(this, "Charset");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "Charset", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSForms 2
		/// Get
		/// </summary>
		[SupportByVersion("MSForms", 2)]
		public virtual Int32 hFont
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteInt32PropertyGet(this, "hFont");
			}
		}

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion MSForms 2
		/// </summary>
		/// <param name="lplpfont">NetOffice.MSFormsApi.IFont lplpfont</param>
		[SupportByVersion("MSForms", 2)]
		public virtual Int32 Clone(out NetOffice.MSFormsApi.IFont lplpfont)
		{
			ParameterModifier[] modifiers = Invoker.CreateParamModifiers(true);
			lplpfont = null;
			object[] paramsArray = Invoker.ValidateParamsArray(new object());
			object returnItem = Invoker.MethodReturn(this, "Clone", paramsArray, modifiers);
            if (paramsArray[1] is MarshalByRefObject)
                paramsArray[1] = Factory.CreateKnownObjectFromComProxy<NetOffice.MSFormsApi.IFont>(this, paramsArray[1], typeof(NetOffice.MSFormsApi.IFont));
            else
                paramsArray[1] = null;
            lplpfont = (NetOffice.MSFormsApi.IFont)paramsArray[0];
			return NetRuntimeSystem.Convert.ToInt32(returnItem);           
        }

		/// <summary>
		/// SupportByVersion MSForms 2
		/// </summary>
		/// <param name="lpFontOther">NetOffice.MSFormsApi.IFont lpFontOther</param>
		[SupportByVersion("MSForms", 2)]
		public virtual Int32 IsEqual(NetOffice.MSFormsApi.IFont lpFontOther)
		{
			return InvokerService.InvokeInternal.ExecuteInt32MethodGet(this, "IsEqual", lpFontOther);
		}

		/// <summary>
		/// SupportByVersion MSForms 2
		/// </summary>
		/// <param name="cyLogical">Int32 cyLogical</param>
		/// <param name="cyHimetric">Int32 cyHimetric</param>
		[SupportByVersion("MSForms", 2)]
		public virtual Int32 SetRatio(Int32 cyLogical, Int32 cyHimetric)
		{
			return InvokerService.InvokeInternal.ExecuteInt32MethodGet(this, "SetRatio", cyLogical, cyHimetric);
		}

		/// <summary>
		/// SupportByVersion MSForms 2
		/// </summary>
		/// <param name="hFont">Int32 hFont</param>
		[SupportByVersion("MSForms", 2)]
		public virtual Int32 AddRefHfont(Int32 hFont)
		{
			return InvokerService.InvokeInternal.ExecuteInt32MethodGet(this, "AddRefHfont", hFont);
		}

		/// <summary>
		/// SupportByVersion MSForms 2
		/// </summary>
		/// <param name="hFont">Int32 hFont</param>
		[SupportByVersion("MSForms", 2)]
		public virtual Int32 ReleaseHfont(Int32 hFont)
		{
			return InvokerService.InvokeInternal.ExecuteInt32MethodGet(this, "ReleaseHfont", hFont);
		}

		#endregion

		#pragma warning restore
	}
}
