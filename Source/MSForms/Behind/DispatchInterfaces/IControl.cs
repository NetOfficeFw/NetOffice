using System.Reflection;
using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;
using NetOffice.MSFormsApi;

namespace NetOffice.MSFormsApi.Behind
{
	/// <summary>
	/// DispatchInterface IControl 
	/// SupportByVersion MSForms, 2
	/// </summary>
	[SupportByVersion("MSForms", 2)]
	[EntityType(EntityType.IsDispatchInterface), BaseType]
 	public class IControl : COMObject, NetOffice.MSFormsApi.IControl
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
                    _contractType = typeof(NetOffice.MSFormsApi.IControl);
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
                    _type = typeof(IControl);
                return _type;
            }
        }
        
        #endregion
        
		#region Ctor

		/// <summary>
		/// Stub Ctor, not indented to use
		/// </summary>
		public IControl() : base()
		{

		}

		#endregion
		
		#region Properties

		/// <summary>
		/// SupportByVersion MSForms 2
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSForms", 2)]
		public virtual bool Cancel
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "Cancel");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "Cancel", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSForms 2
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSForms", 2)]
		public virtual string ControlSource
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "ControlSource");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "ControlSource", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSForms 2
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSForms", 2)]
		public virtual string ControlTipText
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "ControlTipText");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "ControlTipText", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSForms 2
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSForms", 2)]
		public virtual bool Default
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "Default");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "Default", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSForms 2
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSForms", 2)]
		public virtual Single Height
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteSinglePropertyGet(this, "Height");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "Height", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSForms 2
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSForms", 2)]
		public virtual Int32 HelpContextID
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteInt32PropertyGet(this, "HelpContextID");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "HelpContextID", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSForms 2
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSForms", 2)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public virtual bool InSelection
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "InSelection");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "InSelection", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSForms 2
		/// Get
		/// </summary>
		[SupportByVersion("MSForms", 2)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public virtual NetOffice.MSFormsApi.Enums.fmLayoutEffect LayoutEffect
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteEnumPropertyGet<NetOffice.MSFormsApi.Enums.fmLayoutEffect>(this, "LayoutEffect");
			}
		}

		/// <summary>
		/// SupportByVersion MSForms 2
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSForms", 2)]
		public virtual Single Left
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteSinglePropertyGet(this, "Left");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "Left", value);
			}
		}

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
		/// Get
		/// </summary>
		[SupportByVersion("MSForms", 2)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public virtual Single OldHeight
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteSinglePropertyGet(this, "OldHeight");
			}
		}

		/// <summary>
		/// SupportByVersion MSForms 2
		/// Get
		/// </summary>
		[SupportByVersion("MSForms", 2)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public virtual Single OldLeft
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteSinglePropertyGet(this, "OldLeft");
			}
		}

		/// <summary>
		/// SupportByVersion MSForms 2
		/// Get
		/// </summary>
		[SupportByVersion("MSForms", 2)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public virtual Single OldTop
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteSinglePropertyGet(this, "OldTop");
			}
		}

		/// <summary>
		/// SupportByVersion MSForms 2
		/// Get
		/// </summary>
		[SupportByVersion("MSForms", 2)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public virtual Single OldWidth
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteSinglePropertyGet(this, "OldWidth");
			}
		}

		/// <summary>
		/// SupportByVersion MSForms 2
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		[SupportByVersion("MSForms", 2), ProxyResult]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public virtual object Object
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteReferencePropertyGet(this, "Object");
			}
		}

		/// <summary>
		/// SupportByVersion MSForms 2
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		[SupportByVersion("MSForms", 2), ProxyResult]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public virtual object Parent
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteReferencePropertyGet(this, "Parent");
			}
		}

		/// <summary>
		/// SupportByVersion MSForms 2
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSForms", 2)]
		public virtual string RowSource
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "RowSource");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "RowSource", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSForms 2
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSForms", 2)]
		public virtual Int16 RowSourceType
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteInt16PropertyGet(this, "RowSourceType");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "RowSourceType", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSForms 2
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSForms", 2)]
		public virtual Int16 TabIndex
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteInt16PropertyGet(this, "TabIndex");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "TabIndex", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSForms 2
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSForms", 2)]
		public virtual bool TabStop
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "TabStop");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "TabStop", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSForms 2
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSForms", 2)]
		public virtual string Tag
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "Tag");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "Tag", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSForms 2
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSForms", 2)]
		public virtual Single Top
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteSinglePropertyGet(this, "Top");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "Top", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSForms 2
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSForms", 2)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public virtual object BoundValue
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "BoundValue");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteVariantPropertySet(this, "BoundValue", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSForms 2
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSForms", 2)]
		public virtual bool Visible
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "Visible");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "Visible", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSForms 2
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSForms", 2)]
		public virtual Single Width
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteSinglePropertyGet(this, "Width");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "Width", value);
			}
		}

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion MSForms 2
		/// </summary>
		/// <param name="height">Int32 height</param>
		[SupportByVersion("MSForms", 2)]
		public virtual void _SetHeight(Int32 height)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "_SetHeight", height);
		}

		/// <summary>
		/// SupportByVersion MSForms 2
		/// </summary>
		/// <param name="height">Int32 height</param>
		[SupportByVersion("MSForms", 2)]
		public virtual void _GetHeight(out Int32 height)
		{
			ParameterModifier[] modifiers = Invoker.CreateParamModifiers(true);
			height = 0;
			object[] paramsArray = Invoker.ValidateParamsArray(height);
			Invoker.Method(this, "_GetHeight", paramsArray, modifiers);
			height = (Int32)paramsArray[0];
		}

		/// <summary>
		/// SupportByVersion MSForms 2
		/// </summary>
		/// <param name="left">Int32 left</param>
		[SupportByVersion("MSForms", 2)]
		public virtual void _SetLeft(Int32 left)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "_SetLeft", left);
		}

		/// <summary>
		/// SupportByVersion MSForms 2
		/// </summary>
		/// <param name="left">Int32 left</param>
		[SupportByVersion("MSForms", 2)]
		public virtual void _GetLeft(out Int32 left)
		{
			ParameterModifier[] modifiers = Invoker.CreateParamModifiers(true);
			left = 0;
			object[] paramsArray = Invoker.ValidateParamsArray(left);
			Invoker.Method(this, "_GetLeft", paramsArray, modifiers);
			left = (Int32)paramsArray[0];
		}

		/// <summary>
		/// SupportByVersion MSForms 2
		/// </summary>
		/// <param name="oldHeight">Int32 oldHeight</param>
		[SupportByVersion("MSForms", 2)]
		public virtual void _GetOldHeight(out Int32 oldHeight)
		{
			ParameterModifier[] modifiers = Invoker.CreateParamModifiers(true);
			oldHeight = 0;
			object[] paramsArray = Invoker.ValidateParamsArray(oldHeight);
			Invoker.Method(this, "_GetOldHeight", paramsArray, modifiers);
			oldHeight = (Int32)paramsArray[0];
		}

		/// <summary>
		/// SupportByVersion MSForms 2
		/// </summary>
		/// <param name="oldLeft">Int32 oldLeft</param>
		[SupportByVersion("MSForms", 2)]
		public virtual void _GetOldLeft(out Int32 oldLeft)
		{
			ParameterModifier[] modifiers = Invoker.CreateParamModifiers(true);
			oldLeft = 0;
			object[] paramsArray = Invoker.ValidateParamsArray(oldLeft);
			Invoker.Method(this, "_GetOldLeft", paramsArray, modifiers);
			oldLeft = (Int32)paramsArray[0];
		}

		/// <summary>
		/// SupportByVersion MSForms 2
		/// </summary>
		/// <param name="oldTop">Int32 oldTop</param>
		[SupportByVersion("MSForms", 2)]
		public virtual void _GetOldTop(out Int32 oldTop)
		{
			ParameterModifier[] modifiers = Invoker.CreateParamModifiers(true);
			oldTop = 0;
			object[] paramsArray = Invoker.ValidateParamsArray(oldTop);
			Invoker.Method(this, "_GetOldTop", paramsArray, modifiers);
			oldTop = (Int32)paramsArray[0];
		}

		/// <summary>
		/// SupportByVersion MSForms 2
		/// </summary>
		/// <param name="oldWidth">Int32 oldWidth</param>
		[SupportByVersion("MSForms", 2)]
		public virtual void _GetOldWidth(out Int32 oldWidth)
		{
			ParameterModifier[] modifiers = Invoker.CreateParamModifiers(true);
			oldWidth = 0;
			object[] paramsArray = Invoker.ValidateParamsArray(oldWidth);
			Invoker.Method(this, "_GetOldWidth", paramsArray, modifiers);
			oldWidth = (Int32)paramsArray[0];
		}

		/// <summary>
		/// SupportByVersion MSForms 2
		/// </summary>
		/// <param name="top">Int32 top</param>
		[SupportByVersion("MSForms", 2)]
		public virtual void _SetTop(Int32 top)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "_SetTop", top);
		}

		/// <summary>
		/// SupportByVersion MSForms 2
		/// </summary>
		/// <param name="top">Int32 top</param>
		[SupportByVersion("MSForms", 2)]
		public virtual void _GetTop(out Int32 top)
		{
			ParameterModifier[] modifiers = Invoker.CreateParamModifiers(true);
			top = 0;
			object[] paramsArray = Invoker.ValidateParamsArray(top);
			Invoker.Method(this, "_GetTop", paramsArray, modifiers);
			top = (Int32)paramsArray[0];
		}

		/// <summary>
		/// SupportByVersion MSForms 2
		/// </summary>
		/// <param name="width">Int32 width</param>
		[SupportByVersion("MSForms", 2)]
		public virtual void _SetWidth(Int32 width)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "_SetWidth", width);
		}

		/// <summary>
		/// SupportByVersion MSForms 2
		/// </summary>
		/// <param name="width">Int32 width</param>
		[SupportByVersion("MSForms", 2)]
		public virtual void _GetWidth(out Int32 width)
		{
			ParameterModifier[] modifiers = Invoker.CreateParamModifiers(true);
			width = 0;
			object[] paramsArray = Invoker.ValidateParamsArray(width);
			Invoker.Method(this, "_GetWidth", paramsArray, modifiers);
			width = (Int32)paramsArray[0];
		}

		/// <summary>
		/// SupportByVersion MSForms 2
		/// </summary>
		/// <param name="left">optional object left</param>
		/// <param name="top">optional object top</param>
		/// <param name="width">optional object width</param>
		/// <param name="height">optional object height</param>
		/// <param name="layout">optional object layout</param>
		[SupportByVersion("MSForms", 2)]
		public virtual void Move(object left, object top, object width, object height, object layout)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "Move", new object[]{ left, top, width, height, layout });
		}

		/// <summary>
		/// SupportByVersion MSForms 2
		/// </summary>
		[CustomMethod]
		[SupportByVersion("MSForms", 2)]
		public virtual void Move()
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "Move");
		}

		/// <summary>
		/// SupportByVersion MSForms 2
		/// </summary>
		/// <param name="left">optional object left</param>
		[CustomMethod]
		[SupportByVersion("MSForms", 2)]
		public virtual void Move(object left)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "Move", left);
		}

		/// <summary>
		/// SupportByVersion MSForms 2
		/// </summary>
		/// <param name="left">optional object left</param>
		/// <param name="top">optional object top</param>
		[CustomMethod]
		[SupportByVersion("MSForms", 2)]
		public virtual void Move(object left, object top)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "Move", left, top);
		}

		/// <summary>
		/// SupportByVersion MSForms 2
		/// </summary>
		/// <param name="left">optional object left</param>
		/// <param name="top">optional object top</param>
		/// <param name="width">optional object width</param>
		[CustomMethod]
		[SupportByVersion("MSForms", 2)]
		public virtual void Move(object left, object top, object width)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "Move", left, top, width);
		}

		/// <summary>
		/// SupportByVersion MSForms 2
		/// </summary>
		/// <param name="left">optional object left</param>
		/// <param name="top">optional object top</param>
		/// <param name="width">optional object width</param>
		/// <param name="height">optional object height</param>
		[CustomMethod]
		[SupportByVersion("MSForms", 2)]
		public virtual void Move(object left, object top, object width, object height)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "Move", left, top, width, height);
		}

		/// <summary>
		/// SupportByVersion MSForms 2
		/// </summary>
		/// <param name="zPosition">optional object zPosition</param>
		[SupportByVersion("MSForms", 2)]
		public virtual void ZOrder(object zPosition)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "ZOrder", zPosition);
		}

		/// <summary>
		/// SupportByVersion MSForms 2
		/// </summary>
		[CustomMethod]
		[SupportByVersion("MSForms", 2)]
		public virtual void ZOrder()
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "ZOrder");
		}

		/// <summary>
		/// SupportByVersion MSForms 2
		/// </summary>
		/// <param name="selectInGroup">bool selectInGroup</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("MSForms", 2)]
		public virtual void Select(bool selectInGroup)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "Select", selectInGroup);
		}

		/// <summary>
		/// SupportByVersion MSForms 2
		/// </summary>
		[SupportByVersion("MSForms", 2)]
		public virtual void SetFocus()
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "SetFocus");
		}

		/// <summary>
		/// SupportByVersion MSForms 2
		/// </summary>
		[SupportByVersion("MSForms", 2)]
		public virtual Int32 _GethWnd()
		{
			return InvokerService.InvokeInternal.ExecuteInt32MethodGet(this, "_GethWnd");
		}

		/// <summary>
		/// SupportByVersion MSForms 2
		/// </summary>
		[SupportByVersion("MSForms", 2)]
		public virtual Int32 _GetID()
		{
			return InvokerService.InvokeInternal.ExecuteInt32MethodGet(this, "_GetID");
		}

		/// <summary>
		/// SupportByVersion MSForms 2
		/// </summary>
		/// <param name="left">Int32 left</param>
		/// <param name="top">Int32 top</param>
		/// <param name="width">Int32 width</param>
		/// <param name="height">Int32 height</param>
		[SupportByVersion("MSForms", 2)]
		public virtual void _Move(Int32 left, Int32 top, Int32 width, Int32 height)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "_Move", left, top, width, height);
		}

		/// <summary>
		/// SupportByVersion MSForms 2
		/// </summary>
		/// <param name="zPosition">NetOffice.MSFormsApi.Enums.fmZOrder zPosition</param>
		[SupportByVersion("MSForms", 2)]
		public virtual void _ZOrder(NetOffice.MSFormsApi.Enums.fmZOrder zPosition)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "_ZOrder", zPosition);
		}

		#endregion

		#pragma warning restore
	}
}

