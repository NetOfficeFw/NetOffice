using System.Reflection;
using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.MSFormsApi
{
	/// <summary>
	/// DispatchInterface IOptionFrame 
	/// SupportByVersion MSForms, 2
	/// </summary>
	[SupportByVersion("MSForms", 2)]
	[EntityType(EntityType.IsDispatchInterface), BaseType]
 	public class IOptionFrame : COMObject
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
                    _type = typeof(IOptionFrame);
                return _type;
            }
        }
        
        #endregion
        
		#region Ctor

		/// <param name="factory">current used factory core</param>
		/// <param name="parentObject">object there has created the proxy</param>
		/// <param name="proxyShare">proxy share instead if com proxy</param>
		public IOptionFrame(Core factory, ICOMObject parentObject, COMProxyShare proxyShare) : base(factory, parentObject, proxyShare)
		{
		}

		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
		public IOptionFrame(Core factory, ICOMObject parentObject, object comProxy) : base(factory, parentObject, comProxy)
		{
			
		}

        ///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public IOptionFrame(ICOMObject parentObject, object comProxy) : base(parentObject, comProxy)
		{
		}
		
		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public IOptionFrame(Core factory, ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(factory, parentObject, comProxy, comProxyType)
		{

		}

		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public IOptionFrame(ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(parentObject, comProxy, comProxyType)
		{
		}
		
		///<param name="replacedObject">object to replaced. replacedObject are not usable after this action</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public IOptionFrame(ICOMObject replacedObject) : base(replacedObject)
		{
		}
		
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public IOptionFrame() : base()
		{
		}
		
		/// <param name="progId">registered progID</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public IOptionFrame(string progId) : base(progId)
		{
		}
		
		#endregion
		
		#region Properties

		/// <summary>
		/// SupportByVersion MSForms 2
		/// Get
		/// </summary>
		[SupportByVersion("MSForms", 2)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public NetOffice.MSFormsApi.Control ActiveControl
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.MSFormsApi.Control>(this, "ActiveControl", NetOffice.MSFormsApi.Control.LateBindingApiWrapperType);
			}
		}

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
		public Int32 BorderColor
		{
			get
			{
				return Factory.ExecuteInt32PropertyGet(this, "BorderColor");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "BorderColor", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSForms 2
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSForms", 2)]
		public NetOffice.MSFormsApi.Enums.fmBorderStyle BorderStyle
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.MSFormsApi.Enums.fmBorderStyle>(this, "BorderStyle");
			}
			set
			{
				Factory.ExecuteEnumPropertySet(this, "BorderStyle", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSForms 2
		/// Get
		/// </summary>
		[SupportByVersion("MSForms", 2)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public bool CanPaste
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "CanPaste");
			}
		}

		/// <summary>
		/// SupportByVersion MSForms 2
		/// Get
		/// </summary>
		[SupportByVersion("MSForms", 2)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public bool CanRedo
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "CanRedo");
			}
		}

		/// <summary>
		/// SupportByVersion MSForms 2
		/// Get
		/// </summary>
		[SupportByVersion("MSForms", 2)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public bool CanUndo
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "CanUndo");
			}
		}

		/// <summary>
		/// SupportByVersion MSForms 2
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSForms", 2)]
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
		/// SupportByVersion MSForms 2
		/// Get
		/// </summary>
		[SupportByVersion("MSForms", 2)]
		public NetOffice.MSFormsApi.Controls Controls
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.MSFormsApi.Controls>(this, "Controls", NetOffice.MSFormsApi.Controls.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion MSForms 2
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSForms", 2)]
		public NetOffice.MSFormsApi.Enums.fmCycle Cycle
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.MSFormsApi.Enums.fmCycle>(this, "Cycle");
			}
			set
			{
				Factory.ExecuteEnumPropertySet(this, "Cycle", value);
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
		/// Get
		/// </summary>
		[SupportByVersion("MSForms", 2)]
		public Single InsideHeight
		{
			get
			{
				return Factory.ExecuteSinglePropertyGet(this, "InsideHeight");
			}
		}

		/// <summary>
		/// SupportByVersion MSForms 2
		/// Get
		/// </summary>
		[SupportByVersion("MSForms", 2)]
		public Single InsideWidth
		{
			get
			{
				return Factory.ExecuteSinglePropertyGet(this, "InsideWidth");
			}
		}

		/// <summary>
		/// SupportByVersion MSForms 2
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSForms", 2)]
		public NetOffice.MSFormsApi.Enums.fmScrollBars KeepScrollBarsVisible
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.MSFormsApi.Enums.fmScrollBars>(this, "KeepScrollBarsVisible");
			}
			set
			{
				Factory.ExecuteEnumPropertySet(this, "KeepScrollBarsVisible", value);
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
		public NetOffice.MSFormsApi.Enums.fmPictureAlignment PictureAlignment
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.MSFormsApi.Enums.fmPictureAlignment>(this, "PictureAlignment");
			}
			set
			{
				Factory.ExecuteEnumPropertySet(this, "PictureAlignment", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSForms 2
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSForms", 2)]
		public stdole.Picture Picture
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Picture", paramsArray);
                return returnItem as stdole.Picture;
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
		[SupportByVersion("MSForms", 2)]
		public NetOffice.MSFormsApi.Enums.fmPictureSizeMode PictureSizeMode
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.MSFormsApi.Enums.fmPictureSizeMode>(this, "PictureSizeMode");
			}
			set
			{
				Factory.ExecuteEnumPropertySet(this, "PictureSizeMode", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSForms 2
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSForms", 2)]
		public bool PictureTiling
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "PictureTiling");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "PictureTiling", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSForms 2
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSForms", 2)]
		public NetOffice.MSFormsApi.Enums.fmScrollBars ScrollBars
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.MSFormsApi.Enums.fmScrollBars>(this, "ScrollBars");
			}
			set
			{
				Factory.ExecuteEnumPropertySet(this, "ScrollBars", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSForms 2
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSForms", 2)]
		public Single ScrollHeight
		{
			get
			{
				return Factory.ExecuteSinglePropertyGet(this, "ScrollHeight");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "ScrollHeight", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSForms 2
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSForms", 2)]
		public Single ScrollLeft
		{
			get
			{
				return Factory.ExecuteSinglePropertyGet(this, "ScrollLeft");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "ScrollLeft", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSForms 2
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSForms", 2)]
		public Single ScrollTop
		{
			get
			{
				return Factory.ExecuteSinglePropertyGet(this, "ScrollTop");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "ScrollTop", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSForms 2
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSForms", 2)]
		public Single ScrollWidth
		{
			get
			{
				return Factory.ExecuteSinglePropertyGet(this, "ScrollWidth");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "ScrollWidth", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSForms 2
		/// Get
		/// </summary>
		[SupportByVersion("MSForms", 2)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public NetOffice.MSFormsApi.Controls Selected
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.MSFormsApi.Controls>(this, "Selected", NetOffice.MSFormsApi.Controls.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion MSForms 2
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSForms", 2)]
		public NetOffice.MSFormsApi.Enums.fmSpecialEffect SpecialEffect
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.MSFormsApi.Enums.fmSpecialEffect>(this, "SpecialEffect");
			}
			set
			{
				Factory.ExecuteEnumPropertySet(this, "SpecialEffect", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSForms 2
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSForms", 2)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public NetOffice.MSFormsApi.Enums.fmVerticalScrollBarSide VerticalScrollBarSide
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.MSFormsApi.Enums.fmVerticalScrollBarSide>(this, "VerticalScrollBarSide");
			}
			set
			{
				Factory.ExecuteEnumPropertySet(this, "VerticalScrollBarSide", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSForms 2
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSForms", 2)]
		public Int16 Zoom
		{
			get
			{
				return Factory.ExecuteInt16PropertyGet(this, "Zoom");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "Zoom", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSForms 2
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSForms", 2)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public NetOffice.MSFormsApi.Enums.fmMode DesignMode
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.MSFormsApi.Enums.fmMode>(this, "DesignMode");
			}
			set
			{
				Factory.ExecuteEnumPropertySet(this, "DesignMode", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSForms 2
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSForms", 2)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public NetOffice.MSFormsApi.Enums.fmMode ShowToolbox
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.MSFormsApi.Enums.fmMode>(this, "ShowToolbox");
			}
			set
			{
				Factory.ExecuteEnumPropertySet(this, "ShowToolbox", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSForms 2
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSForms", 2)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public NetOffice.MSFormsApi.Enums.fmMode ShowGridDots
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.MSFormsApi.Enums.fmMode>(this, "ShowGridDots");
			}
			set
			{
				Factory.ExecuteEnumPropertySet(this, "ShowGridDots", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSForms 2
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSForms", 2)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public NetOffice.MSFormsApi.Enums.fmMode SnapToGrid
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.MSFormsApi.Enums.fmMode>(this, "SnapToGrid");
			}
			set
			{
				Factory.ExecuteEnumPropertySet(this, "SnapToGrid", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSForms 2
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSForms", 2)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Single GridX
		{
			get
			{
				return Factory.ExecuteSinglePropertyGet(this, "GridX");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "GridX", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSForms 2
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSForms", 2)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Single GridY
		{
			get
			{
				return Factory.ExecuteSinglePropertyGet(this, "GridY");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "GridY", value);
			}
		}

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion MSForms 2
		/// </summary>
		/// <param name="insideHeight">Int32 insideHeight</param>
		[SupportByVersion("MSForms", 2)]
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
		/// </summary>
		/// <param name="insideWidth">Int32 insideWidth</param>
		[SupportByVersion("MSForms", 2)]
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
		/// </summary>
		/// <param name="scrollHeight">Int32 scrollHeight</param>
		[SupportByVersion("MSForms", 2)]
		public void _SetScrollHeight(Int32 scrollHeight)
		{
			 Factory.ExecuteMethod(this, "_SetScrollHeight", scrollHeight);
		}

		/// <summary>
		/// SupportByVersion MSForms 2
		/// </summary>
		/// <param name="scrollHeight">Int32 scrollHeight</param>
		[SupportByVersion("MSForms", 2)]
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
		/// </summary>
		/// <param name="scrollLeft">Int32 scrollLeft</param>
		[SupportByVersion("MSForms", 2)]
		public void _SetScrollLeft(Int32 scrollLeft)
		{
			 Factory.ExecuteMethod(this, "_SetScrollLeft", scrollLeft);
		}

		/// <summary>
		/// SupportByVersion MSForms 2
		/// </summary>
		/// <param name="scrollLeft">Int32 scrollLeft</param>
		[SupportByVersion("MSForms", 2)]
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
		/// </summary>
		/// <param name="scrollTop">Int32 scrollTop</param>
		[SupportByVersion("MSForms", 2)]
		public void _SetScrollTop(Int32 scrollTop)
		{
			 Factory.ExecuteMethod(this, "_SetScrollTop", scrollTop);
		}

		/// <summary>
		/// SupportByVersion MSForms 2
		/// </summary>
		/// <param name="scrollTop">Int32 scrollTop</param>
		[SupportByVersion("MSForms", 2)]
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
		/// </summary>
		/// <param name="scrollWidth">Int32 scrollWidth</param>
		[SupportByVersion("MSForms", 2)]
		public void _SetScrollWidth(Int32 scrollWidth)
		{
			 Factory.ExecuteMethod(this, "_SetScrollWidth", scrollWidth);
		}

		/// <summary>
		/// SupportByVersion MSForms 2
		/// </summary>
		/// <param name="scrollWidth">Int32 scrollWidth</param>
		[SupportByVersion("MSForms", 2)]
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
		/// </summary>
		[SupportByVersion("MSForms", 2)]
		public void Copy()
		{
			 Factory.ExecuteMethod(this, "Copy");
		}

		/// <summary>
		/// SupportByVersion MSForms 2
		/// </summary>
		[SupportByVersion("MSForms", 2)]
		public void Cut()
		{
			 Factory.ExecuteMethod(this, "Cut");
		}

		/// <summary>
		/// SupportByVersion MSForms 2
		/// </summary>
		[SupportByVersion("MSForms", 2)]
		public void Paste()
		{
			 Factory.ExecuteMethod(this, "Paste");
		}

		/// <summary>
		/// SupportByVersion MSForms 2
		/// </summary>
		[SupportByVersion("MSForms", 2)]
		public void RedoAction()
		{
			 Factory.ExecuteMethod(this, "RedoAction");
		}

		/// <summary>
		/// SupportByVersion MSForms 2
		/// </summary>
		[SupportByVersion("MSForms", 2)]
		public void Repaint()
		{
			 Factory.ExecuteMethod(this, "Repaint");
		}

		/// <summary>
		/// SupportByVersion MSForms 2
		/// </summary>
		/// <param name="xAction">optional object xAction</param>
		/// <param name="yAction">optional object yAction</param>
		[SupportByVersion("MSForms", 2)]
		public void Scroll(object xAction, object yAction)
		{
			 Factory.ExecuteMethod(this, "Scroll", xAction, yAction);
		}

		/// <summary>
		/// SupportByVersion MSForms 2
		/// </summary>
		[CustomMethod]
		[SupportByVersion("MSForms", 2)]
		public void Scroll()
		{
			 Factory.ExecuteMethod(this, "Scroll");
		}

		/// <summary>
		/// SupportByVersion MSForms 2
		/// </summary>
		/// <param name="xAction">optional object xAction</param>
		[CustomMethod]
		[SupportByVersion("MSForms", 2)]
		public void Scroll(object xAction)
		{
			 Factory.ExecuteMethod(this, "Scroll", xAction);
		}

		/// <summary>
		/// SupportByVersion MSForms 2
		/// </summary>
		[SupportByVersion("MSForms", 2)]
		public void SetDefaultTabOrder()
		{
			 Factory.ExecuteMethod(this, "SetDefaultTabOrder");
		}

		/// <summary>
		/// SupportByVersion MSForms 2
		/// </summary>
		[SupportByVersion("MSForms", 2)]
		public void UndoAction()
		{
			 Factory.ExecuteMethod(this, "UndoAction");
		}

		/// <summary>
		/// SupportByVersion MSForms 2
		/// </summary>
		/// <param name="gridX">Int32 gridX</param>
		[SupportByVersion("MSForms", 2)]
		public void _SetGridX(Int32 gridX)
		{
			 Factory.ExecuteMethod(this, "_SetGridX", gridX);
		}

		/// <summary>
		/// SupportByVersion MSForms 2
		/// </summary>
		/// <param name="gridX">Int32 gridX</param>
		[SupportByVersion("MSForms", 2)]
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
		/// </summary>
		/// <param name="gridY">Int32 gridY</param>
		[SupportByVersion("MSForms", 2)]
		public void _SetGridY(Int32 gridY)
		{
			 Factory.ExecuteMethod(this, "_SetGridY", gridY);
		}

		/// <summary>
		/// SupportByVersion MSForms 2
		/// </summary>
		/// <param name="gridY">Int32 gridY</param>
		[SupportByVersion("MSForms", 2)]
		public void _GetGridY(out Int32 gridY)
		{
			ParameterModifier[] modifiers = Invoker.CreateParamModifiers(true);
			gridY = 0;
			object[] paramsArray = Invoker.ValidateParamsArray(gridY);
			Invoker.Method(this, "_GetGridY", paramsArray, modifiers);
			gridY = (Int32)paramsArray[0];
		}

		#endregion

		#pragma warning restore
	}
}
