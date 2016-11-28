using System;
using NetRuntimeSystem = System;
using System.Runtime.InteropServices;
using System.Runtime.CompilerServices;
using System.ComponentModel;
using System.Reflection;
using System.Collections.Generic;
using NetOffice;
namespace NetOffice.MSFormsApi
{
	///<summary>
	/// DispatchInterface IControl 
	/// SupportByVersion MSForms, 2
	///</summary>
	[SupportByVersionAttribute("MSForms", 2)]
	[EntityTypeAttribute(EntityType.IsDispatchInterface)]
	public class IControl : COMObject
	{
		#pragma warning disable
		#region Type Information

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
        
		#region Construction

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
		
		/// <param name="progId">registered ProgID</param>
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
		[SupportByVersionAttribute("MSForms", 2)]
		public bool Cancel
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Cancel", paramsArray);
				return NetRuntimeSystem.Convert.ToBoolean(returnItem);
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "Cancel", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion MSForms 2
		/// Get/Set
		/// </summary>
		[SupportByVersionAttribute("MSForms", 2)]
		public string ControlSource
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "ControlSource", paramsArray);
				return NetRuntimeSystem.Convert.ToString(returnItem);
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "ControlSource", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion MSForms 2
		/// Get/Set
		/// </summary>
		[SupportByVersionAttribute("MSForms", 2)]
		public string ControlTipText
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "ControlTipText", paramsArray);
				return NetRuntimeSystem.Convert.ToString(returnItem);
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "ControlTipText", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion MSForms 2
		/// Get/Set
		/// </summary>
		[SupportByVersionAttribute("MSForms", 2)]
		public bool Default
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Default", paramsArray);
				return NetRuntimeSystem.Convert.ToBoolean(returnItem);
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "Default", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion MSForms 2
		/// Get/Set
		/// </summary>
		[SupportByVersionAttribute("MSForms", 2)]
		public Single Height
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Height", paramsArray);
				return NetRuntimeSystem.Convert.ToSingle(returnItem);
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "Height", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion MSForms 2
		/// Get/Set
		/// </summary>
		[SupportByVersionAttribute("MSForms", 2)]
		public Int32 HelpContextID
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "HelpContextID", paramsArray);
				return NetRuntimeSystem.Convert.ToInt32(returnItem);
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "HelpContextID", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion MSForms 2
		/// Get/Set
		/// </summary>
		[SupportByVersionAttribute("MSForms", 2)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public bool InSelection
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "InSelection", paramsArray);
				return NetRuntimeSystem.Convert.ToBoolean(returnItem);
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "InSelection", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion MSForms 2
		/// Get
		/// </summary>
		[SupportByVersionAttribute("MSForms", 2)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public NetOffice.MSFormsApi.Enums.fmLayoutEffect LayoutEffect
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "LayoutEffect", paramsArray);
				int intReturnItem = NetRuntimeSystem.Convert.ToInt32(returnItem);
				return (NetOffice.MSFormsApi.Enums.fmLayoutEffect)intReturnItem;
			}
		}

		/// <summary>
		/// SupportByVersion MSForms 2
		/// Get/Set
		/// </summary>
		[SupportByVersionAttribute("MSForms", 2)]
		public Single Left
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Left", paramsArray);
				return NetRuntimeSystem.Convert.ToSingle(returnItem);
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "Left", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion MSForms 2
		/// Get/Set
		/// </summary>
		[SupportByVersionAttribute("MSForms", 2)]
		public string Name
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Name", paramsArray);
				return NetRuntimeSystem.Convert.ToString(returnItem);
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "Name", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion MSForms 2
		/// Get
		/// </summary>
		[SupportByVersionAttribute("MSForms", 2)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Single OldHeight
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "OldHeight", paramsArray);
				return NetRuntimeSystem.Convert.ToSingle(returnItem);
			}
		}

		/// <summary>
		/// SupportByVersion MSForms 2
		/// Get
		/// </summary>
		[SupportByVersionAttribute("MSForms", 2)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Single OldLeft
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "OldLeft", paramsArray);
				return NetRuntimeSystem.Convert.ToSingle(returnItem);
			}
		}

		/// <summary>
		/// SupportByVersion MSForms 2
		/// Get
		/// </summary>
		[SupportByVersionAttribute("MSForms", 2)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Single OldTop
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "OldTop", paramsArray);
				return NetRuntimeSystem.Convert.ToSingle(returnItem);
			}
		}

		/// <summary>
		/// SupportByVersion MSForms 2
		/// Get
		/// </summary>
		[SupportByVersionAttribute("MSForms", 2)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Single OldWidth
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "OldWidth", paramsArray);
				return NetRuntimeSystem.Convert.ToSingle(returnItem);
			}
		}

		/// <summary>
		/// SupportByVersion MSForms 2
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		[SupportByVersionAttribute("MSForms", 2)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public object Object
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Object", paramsArray);
				ICOMObject newObject = Factory.CreateObjectFromComProxy(this,returnItem);
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion MSForms 2
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		[SupportByVersionAttribute("MSForms", 2)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public object Parent
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Parent", paramsArray);
				ICOMObject newObject = Factory.CreateObjectFromComProxy(this,returnItem);
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion MSForms 2
		/// Get/Set
		/// </summary>
		[SupportByVersionAttribute("MSForms", 2)]
		public string RowSource
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "RowSource", paramsArray);
				return NetRuntimeSystem.Convert.ToString(returnItem);
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "RowSource", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion MSForms 2
		/// Get/Set
		/// </summary>
		[SupportByVersionAttribute("MSForms", 2)]
		public Int16 RowSourceType
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "RowSourceType", paramsArray);
				return NetRuntimeSystem.Convert.ToInt16(returnItem);
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "RowSourceType", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion MSForms 2
		/// Get/Set
		/// </summary>
		[SupportByVersionAttribute("MSForms", 2)]
		public Int16 TabIndex
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "TabIndex", paramsArray);
				return NetRuntimeSystem.Convert.ToInt16(returnItem);
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "TabIndex", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion MSForms 2
		/// Get/Set
		/// </summary>
		[SupportByVersionAttribute("MSForms", 2)]
		public bool TabStop
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "TabStop", paramsArray);
				return NetRuntimeSystem.Convert.ToBoolean(returnItem);
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "TabStop", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion MSForms 2
		/// Get/Set
		/// </summary>
		[SupportByVersionAttribute("MSForms", 2)]
		public string Tag
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Tag", paramsArray);
				return NetRuntimeSystem.Convert.ToString(returnItem);
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "Tag", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion MSForms 2
		/// Get/Set
		/// </summary>
		[SupportByVersionAttribute("MSForms", 2)]
		public Single Top
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Top", paramsArray);
				return NetRuntimeSystem.Convert.ToSingle(returnItem);
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "Top", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion MSForms 2
		/// Get/Set
		/// </summary>
		[SupportByVersionAttribute("MSForms", 2)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public object BoundValue
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "BoundValue", paramsArray);
				if((null != returnItem) && (returnItem is MarshalByRefObject))
				{
					ICOMObject newObject = Factory.CreateObjectFromComProxy(this, returnItem);
					return newObject;
				}
				else
				{
					return  returnItem;
				}
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "BoundValue", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion MSForms 2
		/// Get/Set
		/// </summary>
		[SupportByVersionAttribute("MSForms", 2)]
		public bool Visible
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Visible", paramsArray);
				return NetRuntimeSystem.Convert.ToBoolean(returnItem);
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "Visible", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion MSForms 2
		/// Get/Set
		/// </summary>
		[SupportByVersionAttribute("MSForms", 2)]
		public Single Width
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Width", paramsArray);
				return NetRuntimeSystem.Convert.ToSingle(returnItem);
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "Width", paramsArray);
			}
		}

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion MSForms 2
		/// 
		/// </summary>
		/// <param name="height">Int32 Height</param>
		[SupportByVersionAttribute("MSForms", 2)]
		public void _SetHeight(Int32 height)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(height);
			Invoker.Method(this, "_SetHeight", paramsArray);
		}

		/// <summary>
		/// SupportByVersion MSForms 2
		/// 
		/// </summary>
		/// <param name="height">Int32 Height</param>
		[SupportByVersionAttribute("MSForms", 2)]
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
		/// 
		/// </summary>
		/// <param name="left">Int32 Left</param>
		[SupportByVersionAttribute("MSForms", 2)]
		public void _SetLeft(Int32 left)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(left);
			Invoker.Method(this, "_SetLeft", paramsArray);
		}

		/// <summary>
		/// SupportByVersion MSForms 2
		/// 
		/// </summary>
		/// <param name="left">Int32 Left</param>
		[SupportByVersionAttribute("MSForms", 2)]
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
		/// 
		/// </summary>
		/// <param name="oldHeight">Int32 OldHeight</param>
		[SupportByVersionAttribute("MSForms", 2)]
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
		/// 
		/// </summary>
		/// <param name="oldLeft">Int32 OldLeft</param>
		[SupportByVersionAttribute("MSForms", 2)]
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
		/// 
		/// </summary>
		/// <param name="oldTop">Int32 OldTop</param>
		[SupportByVersionAttribute("MSForms", 2)]
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
		/// 
		/// </summary>
		/// <param name="oldWidth">Int32 OldWidth</param>
		[SupportByVersionAttribute("MSForms", 2)]
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
		/// 
		/// </summary>
		/// <param name="top">Int32 Top</param>
		[SupportByVersionAttribute("MSForms", 2)]
		public void _SetTop(Int32 top)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(top);
			Invoker.Method(this, "_SetTop", paramsArray);
		}

		/// <summary>
		/// SupportByVersion MSForms 2
		/// 
		/// </summary>
		/// <param name="top">Int32 Top</param>
		[SupportByVersionAttribute("MSForms", 2)]
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
		/// 
		/// </summary>
		/// <param name="width">Int32 Width</param>
		[SupportByVersionAttribute("MSForms", 2)]
		public void _SetWidth(Int32 width)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(width);
			Invoker.Method(this, "_SetWidth", paramsArray);
		}

		/// <summary>
		/// SupportByVersion MSForms 2
		/// 
		/// </summary>
		/// <param name="width">Int32 Width</param>
		[SupportByVersionAttribute("MSForms", 2)]
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
		/// 
		/// </summary>
		/// <param name="left">optional object Left</param>
		/// <param name="top">optional object Top</param>
		/// <param name="width">optional object Width</param>
		/// <param name="height">optional object Height</param>
		/// <param name="layout">optional object Layout</param>
		[SupportByVersionAttribute("MSForms", 2)]
		public void Move(object left, object top, object width, object height, object layout)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(left, top, width, height, layout);
			Invoker.Method(this, "Move", paramsArray);
		}

		/// <summary>
		/// SupportByVersion MSForms 2
		/// 
		/// </summary>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("MSForms", 2)]
		public void Move()
		{
			object[] paramsArray = null;
			Invoker.Method(this, "Move", paramsArray);
		}

		/// <summary>
		/// SupportByVersion MSForms 2
		/// 
		/// </summary>
		/// <param name="left">optional object Left</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("MSForms", 2)]
		public void Move(object left)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(left);
			Invoker.Method(this, "Move", paramsArray);
		}

		/// <summary>
		/// SupportByVersion MSForms 2
		/// 
		/// </summary>
		/// <param name="left">optional object Left</param>
		/// <param name="top">optional object Top</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("MSForms", 2)]
		public void Move(object left, object top)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(left, top);
			Invoker.Method(this, "Move", paramsArray);
		}

		/// <summary>
		/// SupportByVersion MSForms 2
		/// 
		/// </summary>
		/// <param name="left">optional object Left</param>
		/// <param name="top">optional object Top</param>
		/// <param name="width">optional object Width</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("MSForms", 2)]
		public void Move(object left, object top, object width)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(left, top, width);
			Invoker.Method(this, "Move", paramsArray);
		}

		/// <summary>
		/// SupportByVersion MSForms 2
		/// 
		/// </summary>
		/// <param name="left">optional object Left</param>
		/// <param name="top">optional object Top</param>
		/// <param name="width">optional object Width</param>
		/// <param name="height">optional object Height</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("MSForms", 2)]
		public void Move(object left, object top, object width, object height)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(left, top, width, height);
			Invoker.Method(this, "Move", paramsArray);
		}

		/// <summary>
		/// SupportByVersion MSForms 2
		/// 
		/// </summary>
		/// <param name="zPosition">optional object zPosition</param>
		[SupportByVersionAttribute("MSForms", 2)]
		public void ZOrder(object zPosition)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(zPosition);
			Invoker.Method(this, "ZOrder", paramsArray);
		}

		/// <summary>
		/// SupportByVersion MSForms 2
		/// 
		/// </summary>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("MSForms", 2)]
		public void ZOrder()
		{
			object[] paramsArray = null;
			Invoker.Method(this, "ZOrder", paramsArray);
		}

		/// <summary>
		/// SupportByVersion MSForms 2
		/// 
		/// </summary>
		/// <param name="selectInGroup">bool SelectInGroup</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersionAttribute("MSForms", 2)]
		public void Select(bool selectInGroup)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(selectInGroup);
			Invoker.Method(this, "Select", paramsArray);
		}

		/// <summary>
		/// SupportByVersion MSForms 2
		/// 
		/// </summary>
		[SupportByVersionAttribute("MSForms", 2)]
		public void SetFocus()
		{
			object[] paramsArray = null;
			Invoker.Method(this, "SetFocus", paramsArray);
		}

		/// <summary>
		/// SupportByVersion MSForms 2
		/// 
		/// </summary>
		[SupportByVersionAttribute("MSForms", 2)]
		public Int32 _GethWnd()
		{
			object[] paramsArray = null;
			object returnItem = Invoker.MethodReturn(this, "_GethWnd", paramsArray);
			return NetRuntimeSystem.Convert.ToInt32(returnItem);
		}

		/// <summary>
		/// SupportByVersion MSForms 2
		/// 
		/// </summary>
		[SupportByVersionAttribute("MSForms", 2)]
		public Int32 _GetID()
		{
			object[] paramsArray = null;
			object returnItem = Invoker.MethodReturn(this, "_GetID", paramsArray);
			return NetRuntimeSystem.Convert.ToInt32(returnItem);
		}

		/// <summary>
		/// SupportByVersion MSForms 2
		/// 
		/// </summary>
		/// <param name="left">Int32 Left</param>
		/// <param name="top">Int32 Top</param>
		/// <param name="width">Int32 Width</param>
		/// <param name="height">Int32 Height</param>
		[SupportByVersionAttribute("MSForms", 2)]
		public void _Move(Int32 left, Int32 top, Int32 width, Int32 height)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(left, top, width, height);
			Invoker.Method(this, "_Move", paramsArray);
		}

		/// <summary>
		/// SupportByVersion MSForms 2
		/// 
		/// </summary>
		/// <param name="zPosition">NetOffice.MSFormsApi.Enums.fmZOrder zPosition</param>
		[SupportByVersionAttribute("MSForms", 2)]
		public void _ZOrder(NetOffice.MSFormsApi.Enums.fmZOrder zPosition)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(zPosition);
			Invoker.Method(this, "_ZOrder", paramsArray);
		}

		#endregion
		#pragma warning restore
	}
}