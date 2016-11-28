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
	/// DispatchInterface IPage 
	/// SupportByVersion MSForms, 2
	///</summary>
	[SupportByVersionAttribute("MSForms", 2)]
	[EntityTypeAttribute(EntityType.IsDispatchInterface)]
	public class IPage : COMObject
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
                    _type = typeof(IPage);
                    
                return _type;
            }
        }
        
        #endregion
        
		#region Construction

		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
		public IPage(Core factory, ICOMObject parentObject, object comProxy) : base(factory, parentObject, comProxy)
		{
			
		}

        ///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public IPage(ICOMObject parentObject, object comProxy) : base(parentObject, comProxy)
		{
		}
		
		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public IPage(Core factory, ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(factory, parentObject, comProxy, comProxyType)
		{

		}

		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public IPage(ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(parentObject, comProxy, comProxyType)
		{
		}
		
		///<param name="replacedObject">object to replaced. replacedObject are not usable after this action</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public IPage(ICOMObject replacedObject) : base(replacedObject)
		{
		}
		
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public IPage() : base()
		{
		}
		
		/// <param name="progId">registered ProgID</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public IPage(string progId) : base(progId)
		{
		}
		
		#endregion
		
		#region Properties

		/// <summary>
		/// SupportByVersion MSForms 2
		/// Get
		/// </summary>
		[SupportByVersionAttribute("MSForms", 2)]
		public NetOffice.MSFormsApi.Controls Controls
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Controls", paramsArray);
				NetOffice.MSFormsApi.Controls newObject = Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.MSFormsApi.Controls.LateBindingApiWrapperType) as NetOffice.MSFormsApi.Controls;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion MSForms 2
		/// Get
		/// </summary>
		[SupportByVersionAttribute("MSForms", 2)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public NetOffice.MSFormsApi.Controls Selected
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Selected", paramsArray);
				NetOffice.MSFormsApi.Controls newObject = Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.MSFormsApi.Controls.LateBindingApiWrapperType) as NetOffice.MSFormsApi.Controls;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion MSForms 2
		/// Get
		/// </summary>
		[SupportByVersionAttribute("MSForms", 2)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public NetOffice.MSFormsApi.Control ActiveControl
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "ActiveControl", paramsArray);
				NetOffice.MSFormsApi.Control newObject = Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.MSFormsApi.Control.LateBindingApiWrapperType) as NetOffice.MSFormsApi.Control;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion MSForms 2
		/// Get
		/// </summary>
		[SupportByVersionAttribute("MSForms", 2)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public bool CanPaste
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "CanPaste", paramsArray);
				return NetRuntimeSystem.Convert.ToBoolean(returnItem);
			}
		}

		/// <summary>
		/// SupportByVersion MSForms 2
		/// Get
		/// </summary>
		[SupportByVersionAttribute("MSForms", 2)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public bool CanRedo
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "CanRedo", paramsArray);
				return NetRuntimeSystem.Convert.ToBoolean(returnItem);
			}
		}

		/// <summary>
		/// SupportByVersion MSForms 2
		/// Get
		/// </summary>
		[SupportByVersionAttribute("MSForms", 2)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public bool CanUndo
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "CanUndo", paramsArray);
				return NetRuntimeSystem.Convert.ToBoolean(returnItem);
			}
		}

		/// <summary>
		/// SupportByVersion MSForms 2
		/// Get/Set
		/// </summary>
		[SupportByVersionAttribute("MSForms", 2)]
		public NetOffice.MSFormsApi.Enums.fmCycle Cycle
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Cycle", paramsArray);
				int intReturnItem = NetRuntimeSystem.Convert.ToInt32(returnItem);
				return (NetOffice.MSFormsApi.Enums.fmCycle)intReturnItem;
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "Cycle", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion MSForms 2
		/// Get/Set
		/// </summary>
		[SupportByVersionAttribute("MSForms", 2)]
		public string Caption
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Caption", paramsArray);
				return NetRuntimeSystem.Convert.ToString(returnItem);
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "Caption", paramsArray);
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
		public bool Enabled
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Enabled", paramsArray);
				return NetRuntimeSystem.Convert.ToBoolean(returnItem);
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "Enabled", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion MSForms 2
		/// Get/Set
		/// </summary>
		[SupportByVersionAttribute("MSForms", 2)]
		public Int32 Index
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Index", paramsArray);
				return NetRuntimeSystem.Convert.ToInt32(returnItem);
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "Index", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion MSForms 2
		/// Get
		/// </summary>
		[SupportByVersionAttribute("MSForms", 2)]
		public Single InsideHeight
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "InsideHeight", paramsArray);
				return NetRuntimeSystem.Convert.ToSingle(returnItem);
			}
		}

		/// <summary>
		/// SupportByVersion MSForms 2
		/// Get
		/// </summary>
		[SupportByVersionAttribute("MSForms", 2)]
		public Single InsideWidth
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "InsideWidth", paramsArray);
				return NetRuntimeSystem.Convert.ToSingle(returnItem);
			}
		}

		/// <summary>
		/// SupportByVersion MSForms 2
		/// Get/Set
		/// </summary>
		[SupportByVersionAttribute("MSForms", 2)]
		public NetOffice.MSFormsApi.Enums.fmScrollBars KeepScrollBarsVisible
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "KeepScrollBarsVisible", paramsArray);
				int intReturnItem = NetRuntimeSystem.Convert.ToInt32(returnItem);
				return (NetOffice.MSFormsApi.Enums.fmScrollBars)intReturnItem;
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "KeepScrollBarsVisible", paramsArray);
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
		/// Get/Set
		/// </summary>
		[SupportByVersionAttribute("MSForms", 2)]
		public NetOffice.MSFormsApi.Enums.fmPictureAlignment PictureAlignment
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "PictureAlignment", paramsArray);
				int intReturnItem = NetRuntimeSystem.Convert.ToInt32(returnItem);
				return (NetOffice.MSFormsApi.Enums.fmPictureAlignment)intReturnItem;
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "PictureAlignment", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion MSForms 2
		/// Get/Set
		/// </summary>
		[SupportByVersionAttribute("MSForms", 2)]
		public stdole.Picture Picture
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Picture", paramsArray);
				stdole.Picture newObject = Factory.CreateObjectFromComProxy(this,returnItem) as stdole.Picture;
				return newObject;
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "Picture", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion MSForms 2
		/// Get/Set
		/// </summary>
		[SupportByVersionAttribute("MSForms", 2)]
		public NetOffice.MSFormsApi.Enums.fmPictureSizeMode PictureSizeMode
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "PictureSizeMode", paramsArray);
				int intReturnItem = NetRuntimeSystem.Convert.ToInt32(returnItem);
				return (NetOffice.MSFormsApi.Enums.fmPictureSizeMode)intReturnItem;
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "PictureSizeMode", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion MSForms 2
		/// Get/Set
		/// </summary>
		[SupportByVersionAttribute("MSForms", 2)]
		public bool PictureTiling
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "PictureTiling", paramsArray);
				return NetRuntimeSystem.Convert.ToBoolean(returnItem);
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "PictureTiling", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion MSForms 2
		/// Get/Set
		/// </summary>
		[SupportByVersionAttribute("MSForms", 2)]
		public NetOffice.MSFormsApi.Enums.fmScrollBars ScrollBars
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "ScrollBars", paramsArray);
				int intReturnItem = NetRuntimeSystem.Convert.ToInt32(returnItem);
				return (NetOffice.MSFormsApi.Enums.fmScrollBars)intReturnItem;
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "ScrollBars", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion MSForms 2
		/// Get/Set
		/// </summary>
		[SupportByVersionAttribute("MSForms", 2)]
		public Single ScrollHeight
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "ScrollHeight", paramsArray);
				return NetRuntimeSystem.Convert.ToSingle(returnItem);
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "ScrollHeight", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion MSForms 2
		/// Get/Set
		/// </summary>
		[SupportByVersionAttribute("MSForms", 2)]
		public Single ScrollLeft
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "ScrollLeft", paramsArray);
				return NetRuntimeSystem.Convert.ToSingle(returnItem);
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "ScrollLeft", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion MSForms 2
		/// Get/Set
		/// </summary>
		[SupportByVersionAttribute("MSForms", 2)]
		public Single ScrollTop
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "ScrollTop", paramsArray);
				return NetRuntimeSystem.Convert.ToSingle(returnItem);
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "ScrollTop", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion MSForms 2
		/// Get/Set
		/// </summary>
		[SupportByVersionAttribute("MSForms", 2)]
		public Single ScrollWidth
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "ScrollWidth", paramsArray);
				return NetRuntimeSystem.Convert.ToSingle(returnItem);
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "ScrollWidth", paramsArray);
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
		public NetOffice.MSFormsApi.Enums.fmTransitionEffect TransitionEffect
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "TransitionEffect", paramsArray);
				int intReturnItem = NetRuntimeSystem.Convert.ToInt32(returnItem);
				return (NetOffice.MSFormsApi.Enums.fmTransitionEffect)intReturnItem;
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "TransitionEffect", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion MSForms 2
		/// Get/Set
		/// </summary>
		[SupportByVersionAttribute("MSForms", 2)]
		public Int32 TransitionPeriod
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "TransitionPeriod", paramsArray);
				return NetRuntimeSystem.Convert.ToInt32(returnItem);
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "TransitionPeriod", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion MSForms 2
		/// Get/Set
		/// </summary>
		[SupportByVersionAttribute("MSForms", 2)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public NetOffice.MSFormsApi.Enums.fmVerticalScrollBarSide VerticalScrollBarSide
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "VerticalScrollBarSide", paramsArray);
				int intReturnItem = NetRuntimeSystem.Convert.ToInt32(returnItem);
				return (NetOffice.MSFormsApi.Enums.fmVerticalScrollBarSide)intReturnItem;
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "VerticalScrollBarSide", paramsArray);
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
		public Int16 Zoom
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Zoom", paramsArray);
				return NetRuntimeSystem.Convert.ToInt16(returnItem);
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "Zoom", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion MSForms 2
		/// Get/Set
		/// </summary>
		[SupportByVersionAttribute("MSForms", 2)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public NetOffice.MSFormsApi.Enums.fmMode DesignMode
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "DesignMode", paramsArray);
				int intReturnItem = NetRuntimeSystem.Convert.ToInt32(returnItem);
				return (NetOffice.MSFormsApi.Enums.fmMode)intReturnItem;
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "DesignMode", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion MSForms 2
		/// Get/Set
		/// </summary>
		[SupportByVersionAttribute("MSForms", 2)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public NetOffice.MSFormsApi.Enums.fmMode ShowToolbox
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "ShowToolbox", paramsArray);
				int intReturnItem = NetRuntimeSystem.Convert.ToInt32(returnItem);
				return (NetOffice.MSFormsApi.Enums.fmMode)intReturnItem;
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "ShowToolbox", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion MSForms 2
		/// Get/Set
		/// </summary>
		[SupportByVersionAttribute("MSForms", 2)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public NetOffice.MSFormsApi.Enums.fmMode ShowGridDots
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "ShowGridDots", paramsArray);
				int intReturnItem = NetRuntimeSystem.Convert.ToInt32(returnItem);
				return (NetOffice.MSFormsApi.Enums.fmMode)intReturnItem;
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "ShowGridDots", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion MSForms 2
		/// Get/Set
		/// </summary>
		[SupportByVersionAttribute("MSForms", 2)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public NetOffice.MSFormsApi.Enums.fmMode SnapToGrid
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "SnapToGrid", paramsArray);
				int intReturnItem = NetRuntimeSystem.Convert.ToInt32(returnItem);
				return (NetOffice.MSFormsApi.Enums.fmMode)intReturnItem;
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "SnapToGrid", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion MSForms 2
		/// Get/Set
		/// </summary>
		[SupportByVersionAttribute("MSForms", 2)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Single GridX
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "GridX", paramsArray);
				return NetRuntimeSystem.Convert.ToSingle(returnItem);
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "GridX", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion MSForms 2
		/// Get/Set
		/// </summary>
		[SupportByVersionAttribute("MSForms", 2)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Single GridY
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "GridY", paramsArray);
				return NetRuntimeSystem.Convert.ToSingle(returnItem);
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "GridY", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion MSForms 2
		/// Get/Set
		/// </summary>
		[SupportByVersionAttribute("MSForms", 2)]
		public string Accelerator
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Accelerator", paramsArray);
				return NetRuntimeSystem.Convert.ToString(returnItem);
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "Accelerator", paramsArray);
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

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion MSForms 2
		/// 
		/// </summary>
		/// <param name="insideHeight">Int32 InsideHeight</param>
		[SupportByVersionAttribute("MSForms", 2)]
		public void _GetInsideHeight(out Int32 insideHeight)
		{
			ParameterModifier[] modifiers = Invoker.CreateParamModifiers(true);
			insideHeight = 0;
			object[] paramsArray = Invoker.ValidateParamsArray(insideHeight);
			Invoker.Method(this, "_GetInsideHeight", paramsArray, modifiers);
			insideHeight = (Int32)paramsArray[0];
		}

		/// <summary>
		/// SupportByVersion MSForms 2
		/// 
		/// </summary>
		/// <param name="insideWidth">Int32 InsideWidth</param>
		[SupportByVersionAttribute("MSForms", 2)]
		public void _GetInsideWidth(out Int32 insideWidth)
		{
			ParameterModifier[] modifiers = Invoker.CreateParamModifiers(true);
			insideWidth = 0;
			object[] paramsArray = Invoker.ValidateParamsArray(insideWidth);
			Invoker.Method(this, "_GetInsideWidth", paramsArray, modifiers);
			insideWidth = (Int32)paramsArray[0];
		}

		/// <summary>
		/// SupportByVersion MSForms 2
		/// 
		/// </summary>
		/// <param name="scrollHeight">Int32 ScrollHeight</param>
		[SupportByVersionAttribute("MSForms", 2)]
		public void _SetScrollHeight(Int32 scrollHeight)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(scrollHeight);
			Invoker.Method(this, "_SetScrollHeight", paramsArray);
		}

		/// <summary>
		/// SupportByVersion MSForms 2
		/// 
		/// </summary>
		/// <param name="scrollHeight">Int32 ScrollHeight</param>
		[SupportByVersionAttribute("MSForms", 2)]
		public void _GetScrollHeight(out Int32 scrollHeight)
		{
			ParameterModifier[] modifiers = Invoker.CreateParamModifiers(true);
			scrollHeight = 0;
			object[] paramsArray = Invoker.ValidateParamsArray(scrollHeight);
			Invoker.Method(this, "_GetScrollHeight", paramsArray, modifiers);
			scrollHeight = (Int32)paramsArray[0];
		}

		/// <summary>
		/// SupportByVersion MSForms 2
		/// 
		/// </summary>
		/// <param name="scrollLeft">Int32 ScrollLeft</param>
		[SupportByVersionAttribute("MSForms", 2)]
		public void _SetScrollLeft(Int32 scrollLeft)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(scrollLeft);
			Invoker.Method(this, "_SetScrollLeft", paramsArray);
		}

		/// <summary>
		/// SupportByVersion MSForms 2
		/// 
		/// </summary>
		/// <param name="scrollLeft">Int32 ScrollLeft</param>
		[SupportByVersionAttribute("MSForms", 2)]
		public void _GetScrollLeft(out Int32 scrollLeft)
		{
			ParameterModifier[] modifiers = Invoker.CreateParamModifiers(true);
			scrollLeft = 0;
			object[] paramsArray = Invoker.ValidateParamsArray(scrollLeft);
			Invoker.Method(this, "_GetScrollLeft", paramsArray, modifiers);
			scrollLeft = (Int32)paramsArray[0];
		}

		/// <summary>
		/// SupportByVersion MSForms 2
		/// 
		/// </summary>
		/// <param name="scrollTop">Int32 ScrollTop</param>
		[SupportByVersionAttribute("MSForms", 2)]
		public void _SetScrollTop(Int32 scrollTop)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(scrollTop);
			Invoker.Method(this, "_SetScrollTop", paramsArray);
		}

		/// <summary>
		/// SupportByVersion MSForms 2
		/// 
		/// </summary>
		/// <param name="scrollTop">Int32 ScrollTop</param>
		[SupportByVersionAttribute("MSForms", 2)]
		public void _GetScrollTop(out Int32 scrollTop)
		{
			ParameterModifier[] modifiers = Invoker.CreateParamModifiers(true);
			scrollTop = 0;
			object[] paramsArray = Invoker.ValidateParamsArray(scrollTop);
			Invoker.Method(this, "_GetScrollTop", paramsArray, modifiers);
			scrollTop = (Int32)paramsArray[0];
		}

		/// <summary>
		/// SupportByVersion MSForms 2
		/// 
		/// </summary>
		/// <param name="scrollWidth">Int32 ScrollWidth</param>
		[SupportByVersionAttribute("MSForms", 2)]
		public void _SetScrollWidth(Int32 scrollWidth)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(scrollWidth);
			Invoker.Method(this, "_SetScrollWidth", paramsArray);
		}

		/// <summary>
		/// SupportByVersion MSForms 2
		/// 
		/// </summary>
		/// <param name="scrollWidth">Int32 ScrollWidth</param>
		[SupportByVersionAttribute("MSForms", 2)]
		public void _GetScrollWidth(out Int32 scrollWidth)
		{
			ParameterModifier[] modifiers = Invoker.CreateParamModifiers(true);
			scrollWidth = 0;
			object[] paramsArray = Invoker.ValidateParamsArray(scrollWidth);
			Invoker.Method(this, "_GetScrollWidth", paramsArray, modifiers);
			scrollWidth = (Int32)paramsArray[0];
		}

		/// <summary>
		/// SupportByVersion MSForms 2
		/// 
		/// </summary>
		/// <param name="gridX">Int32 GridX</param>
		[SupportByVersionAttribute("MSForms", 2)]
		public void _SetGridX(Int32 gridX)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(gridX);
			Invoker.Method(this, "_SetGridX", paramsArray);
		}

		/// <summary>
		/// SupportByVersion MSForms 2
		/// 
		/// </summary>
		/// <param name="gridX">Int32 GridX</param>
		[SupportByVersionAttribute("MSForms", 2)]
		public void _GetGridX(out Int32 gridX)
		{
			ParameterModifier[] modifiers = Invoker.CreateParamModifiers(true);
			gridX = 0;
			object[] paramsArray = Invoker.ValidateParamsArray(gridX);
			Invoker.Method(this, "_GetGridX", paramsArray, modifiers);
			gridX = (Int32)paramsArray[0];
		}

		/// <summary>
		/// SupportByVersion MSForms 2
		/// 
		/// </summary>
		/// <param name="gridY">Int32 GridY</param>
		[SupportByVersionAttribute("MSForms", 2)]
		public void _SetGridY(Int32 gridY)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(gridY);
			Invoker.Method(this, "_SetGridY", paramsArray);
		}

		/// <summary>
		/// SupportByVersion MSForms 2
		/// 
		/// </summary>
		/// <param name="gridY">Int32 GridY</param>
		[SupportByVersionAttribute("MSForms", 2)]
		public void _GetGridY(out Int32 gridY)
		{
			ParameterModifier[] modifiers = Invoker.CreateParamModifiers(true);
			gridY = 0;
			object[] paramsArray = Invoker.ValidateParamsArray(gridY);
			Invoker.Method(this, "_GetGridY", paramsArray, modifiers);
			gridY = (Int32)paramsArray[0];
		}

		/// <summary>
		/// SupportByVersion MSForms 2
		/// 
		/// </summary>
		[SupportByVersionAttribute("MSForms", 2)]
		public void Copy()
		{
			object[] paramsArray = null;
			Invoker.Method(this, "Copy", paramsArray);
		}

		/// <summary>
		/// SupportByVersion MSForms 2
		/// 
		/// </summary>
		[SupportByVersionAttribute("MSForms", 2)]
		public void Cut()
		{
			object[] paramsArray = null;
			Invoker.Method(this, "Cut", paramsArray);
		}

		/// <summary>
		/// SupportByVersion MSForms 2
		/// 
		/// </summary>
		[SupportByVersionAttribute("MSForms", 2)]
		public void Paste()
		{
			object[] paramsArray = null;
			Invoker.Method(this, "Paste", paramsArray);
		}

		/// <summary>
		/// SupportByVersion MSForms 2
		/// 
		/// </summary>
		[SupportByVersionAttribute("MSForms", 2)]
		public void RedoAction()
		{
			object[] paramsArray = null;
			Invoker.Method(this, "RedoAction", paramsArray);
		}

		/// <summary>
		/// SupportByVersion MSForms 2
		/// 
		/// </summary>
		[SupportByVersionAttribute("MSForms", 2)]
		public void Repaint()
		{
			object[] paramsArray = null;
			Invoker.Method(this, "Repaint", paramsArray);
		}

		/// <summary>
		/// SupportByVersion MSForms 2
		/// 
		/// </summary>
		/// <param name="xAction">optional object xAction</param>
		/// <param name="yAction">optional object yAction</param>
		[SupportByVersionAttribute("MSForms", 2)]
		public void Scroll(object xAction, object yAction)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(xAction, yAction);
			Invoker.Method(this, "Scroll", paramsArray);
		}

		/// <summary>
		/// SupportByVersion MSForms 2
		/// 
		/// </summary>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("MSForms", 2)]
		public void Scroll()
		{
			object[] paramsArray = null;
			Invoker.Method(this, "Scroll", paramsArray);
		}

		/// <summary>
		/// SupportByVersion MSForms 2
		/// 
		/// </summary>
		/// <param name="xAction">optional object xAction</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("MSForms", 2)]
		public void Scroll(object xAction)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(xAction);
			Invoker.Method(this, "Scroll", paramsArray);
		}

		/// <summary>
		/// SupportByVersion MSForms 2
		/// 
		/// </summary>
		[SupportByVersionAttribute("MSForms", 2)]
		public void SetDefaultTabOrder()
		{
			object[] paramsArray = null;
			Invoker.Method(this, "SetDefaultTabOrder", paramsArray);
		}

		/// <summary>
		/// SupportByVersion MSForms 2
		/// 
		/// </summary>
		[SupportByVersionAttribute("MSForms", 2)]
		public void UndoAction()
		{
			object[] paramsArray = null;
			Invoker.Method(this, "UndoAction", paramsArray);
		}

		#endregion
		#pragma warning restore
	}
}