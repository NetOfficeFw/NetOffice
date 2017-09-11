using System.Reflection;
using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.MSFormsApi
{
	/// <summary>
	/// DispatchInterface IControl 
	/// SupportByVersion MSForms, 2
	/// </summary>
	[SupportByVersion("MSForms", 2)]
	[EntityType(EntityType.IsDispatchInterface), BaseType]
 	public class IControl : COMObject
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
                    _type = typeof(IControl);
                return _type;
            }
        }
        
        #endregion
        
		#region Ctor

		/// <param name="factory">current used factory core</param>
		/// <param name="parentObject">object there has created the proxy</param>
		/// <param name="proxyShare">proxy share instead if com proxy</param>
		public IControl(Core factory, ICOMObject parentObject, COMProxyShare proxyShare) : base(factory, parentObject, proxyShare)
		{
		}

		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
		public IControl(Core factory, ICOMObject parentObject, object comProxy) : base(factory, parentObject, comProxy)
		{
			
		}

        ///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public IControl(ICOMObject parentObject, object comProxy) : base(parentObject, comProxy)
		{
		}
		
		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public IControl(Core factory, ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(factory, parentObject, comProxy, comProxyType)
		{

		}

		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public IControl(ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(parentObject, comProxy, comProxyType)
		{
		}
		
		///<param name="replacedObject">object to replaced. replacedObject are not usable after this action</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public IControl(ICOMObject replacedObject) : base(replacedObject)
		{
		}
		
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public IControl() : base()
		{
		}
		
		/// <param name="progId">registered progID</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public IControl(string progId) : base(progId)
		{
		}
		
		#endregion
		
		#region Properties

		/// <summary>
		/// SupportByVersion MSForms 2
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSForms", 2)]
		public bool Cancel
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "Cancel");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "Cancel", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSForms 2
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSForms", 2)]
		public string ControlSource
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "ControlSource");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "ControlSource", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSForms 2
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSForms", 2)]
		public string ControlTipText
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "ControlTipText");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "ControlTipText", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSForms 2
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSForms", 2)]
		public bool Default
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "Default");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "Default", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSForms 2
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSForms", 2)]
		public Single Height
		{
			get
			{
				return Factory.ExecuteSinglePropertyGet(this, "Height");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "Height", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSForms 2
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSForms", 2)]
		public Int32 HelpContextID
		{
			get
			{
				return Factory.ExecuteInt32PropertyGet(this, "HelpContextID");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "HelpContextID", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSForms 2
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSForms", 2)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public bool InSelection
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "InSelection");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "InSelection", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSForms 2
		/// Get
		/// </summary>
		[SupportByVersion("MSForms", 2)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public NetOffice.MSFormsApi.Enums.fmLayoutEffect LayoutEffect
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.MSFormsApi.Enums.fmLayoutEffect>(this, "LayoutEffect");
			}
		}

		/// <summary>
		/// SupportByVersion MSForms 2
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSForms", 2)]
		public Single Left
		{
			get
			{
				return Factory.ExecuteSinglePropertyGet(this, "Left");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "Left", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSForms 2
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSForms", 2)]
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
		/// SupportByVersion MSForms 2
		/// Get
		/// </summary>
		[SupportByVersion("MSForms", 2)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Single OldHeight
		{
			get
			{
				return Factory.ExecuteSinglePropertyGet(this, "OldHeight");
			}
		}

		/// <summary>
		/// SupportByVersion MSForms 2
		/// Get
		/// </summary>
		[SupportByVersion("MSForms", 2)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Single OldLeft
		{
			get
			{
				return Factory.ExecuteSinglePropertyGet(this, "OldLeft");
			}
		}

		/// <summary>
		/// SupportByVersion MSForms 2
		/// Get
		/// </summary>
		[SupportByVersion("MSForms", 2)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Single OldTop
		{
			get
			{
				return Factory.ExecuteSinglePropertyGet(this, "OldTop");
			}
		}

		/// <summary>
		/// SupportByVersion MSForms 2
		/// Get
		/// </summary>
		[SupportByVersion("MSForms", 2)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Single OldWidth
		{
			get
			{
				return Factory.ExecuteSinglePropertyGet(this, "OldWidth");
			}
		}

		/// <summary>
		/// SupportByVersion MSForms 2
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		[SupportByVersion("MSForms", 2), ProxyResult]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public object Object
		{
			get
			{
				return Factory.ExecuteReferencePropertyGet(this, "Object");
			}
		}

		/// <summary>
		/// SupportByVersion MSForms 2
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		[SupportByVersion("MSForms", 2), ProxyResult]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public object Parent
		{
			get
			{
				return Factory.ExecuteReferencePropertyGet(this, "Parent");
			}
		}

		/// <summary>
		/// SupportByVersion MSForms 2
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSForms", 2)]
		public string RowSource
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "RowSource");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "RowSource", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSForms 2
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSForms", 2)]
		public Int16 RowSourceType
		{
			get
			{
				return Factory.ExecuteInt16PropertyGet(this, "RowSourceType");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "RowSourceType", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSForms 2
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSForms", 2)]
		public Int16 TabIndex
		{
			get
			{
				return Factory.ExecuteInt16PropertyGet(this, "TabIndex");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "TabIndex", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSForms 2
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSForms", 2)]
		public bool TabStop
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "TabStop");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "TabStop", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSForms 2
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSForms", 2)]
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
		/// SupportByVersion MSForms 2
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSForms", 2)]
		public Single Top
		{
			get
			{
				return Factory.ExecuteSinglePropertyGet(this, "Top");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "Top", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSForms 2
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSForms", 2)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public object BoundValue
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "BoundValue");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "BoundValue", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSForms 2
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSForms", 2)]
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
		/// SupportByVersion MSForms 2
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSForms", 2)]
		public Single Width
		{
			get
			{
				return Factory.ExecuteSinglePropertyGet(this, "Width");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "Width", value);
			}
		}

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion MSForms 2
		/// </summary>
		/// <param name="height">Int32 height</param>
		[SupportByVersion("MSForms", 2)]
		public void _SetHeight(Int32 height)
		{
			 Factory.ExecuteMethod(this, "_SetHeight", height);
		}

		/// <summary>
		/// SupportByVersion MSForms 2
		/// </summary>
		/// <param name="height">Int32 height</param>
		[SupportByVersion("MSForms", 2)]
		public void _GetHeight(out Int32 height)
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
		public void _SetLeft(Int32 left)
		{
			 Factory.ExecuteMethod(this, "_SetLeft", left);
		}

		/// <summary>
		/// SupportByVersion MSForms 2
		/// </summary>
		/// <param name="left">Int32 left</param>
		[SupportByVersion("MSForms", 2)]
		public void _GetLeft(out Int32 left)
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
		public void _GetOldHeight(out Int32 oldHeight)
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
		public void _GetOldLeft(out Int32 oldLeft)
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
		public void _GetOldTop(out Int32 oldTop)
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
		public void _GetOldWidth(out Int32 oldWidth)
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
		public void _SetTop(Int32 top)
		{
			 Factory.ExecuteMethod(this, "_SetTop", top);
		}

		/// <summary>
		/// SupportByVersion MSForms 2
		/// </summary>
		/// <param name="top">Int32 top</param>
		[SupportByVersion("MSForms", 2)]
		public void _GetTop(out Int32 top)
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
		public void _SetWidth(Int32 width)
		{
			 Factory.ExecuteMethod(this, "_SetWidth", width);
		}

		/// <summary>
		/// SupportByVersion MSForms 2
		/// </summary>
		/// <param name="width">Int32 width</param>
		[SupportByVersion("MSForms", 2)]
		public void _GetWidth(out Int32 width)
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
		public void Move(object left, object top, object width, object height, object layout)
		{
			 Factory.ExecuteMethod(this, "Move", new object[]{ left, top, width, height, layout });
		}

		/// <summary>
		/// SupportByVersion MSForms 2
		/// </summary>
		[CustomMethod]
		[SupportByVersion("MSForms", 2)]
		public void Move()
		{
			 Factory.ExecuteMethod(this, "Move");
		}

		/// <summary>
		/// SupportByVersion MSForms 2
		/// </summary>
		/// <param name="left">optional object left</param>
		[CustomMethod]
		[SupportByVersion("MSForms", 2)]
		public void Move(object left)
		{
			 Factory.ExecuteMethod(this, "Move", left);
		}

		/// <summary>
		/// SupportByVersion MSForms 2
		/// </summary>
		/// <param name="left">optional object left</param>
		/// <param name="top">optional object top</param>
		[CustomMethod]
		[SupportByVersion("MSForms", 2)]
		public void Move(object left, object top)
		{
			 Factory.ExecuteMethod(this, "Move", left, top);
		}

		/// <summary>
		/// SupportByVersion MSForms 2
		/// </summary>
		/// <param name="left">optional object left</param>
		/// <param name="top">optional object top</param>
		/// <param name="width">optional object width</param>
		[CustomMethod]
		[SupportByVersion("MSForms", 2)]
		public void Move(object left, object top, object width)
		{
			 Factory.ExecuteMethod(this, "Move", left, top, width);
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
		public void Move(object left, object top, object width, object height)
		{
			 Factory.ExecuteMethod(this, "Move", left, top, width, height);
		}

		/// <summary>
		/// SupportByVersion MSForms 2
		/// </summary>
		/// <param name="zPosition">optional object zPosition</param>
		[SupportByVersion("MSForms", 2)]
		public void ZOrder(object zPosition)
		{
			 Factory.ExecuteMethod(this, "ZOrder", zPosition);
		}

		/// <summary>
		/// SupportByVersion MSForms 2
		/// </summary>
		[CustomMethod]
		[SupportByVersion("MSForms", 2)]
		public void ZOrder()
		{
			 Factory.ExecuteMethod(this, "ZOrder");
		}

		/// <summary>
		/// SupportByVersion MSForms 2
		/// </summary>
		/// <param name="selectInGroup">bool selectInGroup</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("MSForms", 2)]
		public void Select(bool selectInGroup)
		{
			 Factory.ExecuteMethod(this, "Select", selectInGroup);
		}

		/// <summary>
		/// SupportByVersion MSForms 2
		/// </summary>
		[SupportByVersion("MSForms", 2)]
		public void SetFocus()
		{
			 Factory.ExecuteMethod(this, "SetFocus");
		}

		/// <summary>
		/// SupportByVersion MSForms 2
		/// </summary>
		[SupportByVersion("MSForms", 2)]
		public Int32 _GethWnd()
		{
			return Factory.ExecuteInt32MethodGet(this, "_GethWnd");
		}

		/// <summary>
		/// SupportByVersion MSForms 2
		/// </summary>
		[SupportByVersion("MSForms", 2)]
		public Int32 _GetID()
		{
			return Factory.ExecuteInt32MethodGet(this, "_GetID");
		}

		/// <summary>
		/// SupportByVersion MSForms 2
		/// </summary>
		/// <param name="left">Int32 left</param>
		/// <param name="top">Int32 top</param>
		/// <param name="width">Int32 width</param>
		/// <param name="height">Int32 height</param>
		[SupportByVersion("MSForms", 2)]
		public void _Move(Int32 left, Int32 top, Int32 width, Int32 height)
		{
			 Factory.ExecuteMethod(this, "_Move", left, top, width, height);
		}

		/// <summary>
		/// SupportByVersion MSForms 2
		/// </summary>
		/// <param name="zPosition">NetOffice.MSFormsApi.Enums.fmZOrder zPosition</param>
		[SupportByVersion("MSForms", 2)]
		public void _ZOrder(NetOffice.MSFormsApi.Enums.fmZOrder zPosition)
		{
			 Factory.ExecuteMethod(this, "_ZOrder", zPosition);
		}

		#endregion

		#pragma warning restore
	}
}
