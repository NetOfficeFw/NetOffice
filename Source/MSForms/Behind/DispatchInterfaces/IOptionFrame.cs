using System.Reflection;
using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;
using NetOffice.MSFormsApi;

namespace NetOffice.MSFormsApi.Behind
{
	/// <summary>
	/// DispatchInterface IOptionFrame 
	/// SupportByVersion MSForms, 2
	/// </summary>
	[SupportByVersion("MSForms", 2)]
	[EntityType(EntityType.IsDispatchInterface), BaseType]
 	public class IOptionFrame : COMObject, NetOffice.MSFormsApi.IOptionFrame
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
                    _contractType = typeof(NetOffice.MSFormsApi.IOptionFrame);
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
                    _type = typeof(IOptionFrame);
                return _type;
            }
        }
        
        #endregion
        
		#region Ctor

		/// <summary>
		/// Stub Ctor, not indented to use
		/// </summary>
		public IOptionFrame() : base()
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
		public virtual NetOffice.MSFormsApi.Control ActiveControl
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.MSFormsApi.Control>(this, "ActiveControl", typeof(NetOffice.MSFormsApi.Control));
			}
		}

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
		public virtual Int32 BorderColor
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteInt32PropertyGet(this, "BorderColor");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "BorderColor", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSForms 2
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSForms", 2)]
		public virtual NetOffice.MSFormsApi.Enums.fmBorderStyle BorderStyle
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteEnumPropertyGet<NetOffice.MSFormsApi.Enums.fmBorderStyle>(this, "BorderStyle");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteEnumPropertySet(this, "BorderStyle", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSForms 2
		/// Get
		/// </summary>
		[SupportByVersion("MSForms", 2)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public virtual bool CanPaste
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "CanPaste");
			}
		}

		/// <summary>
		/// SupportByVersion MSForms 2
		/// Get
		/// </summary>
		[SupportByVersion("MSForms", 2)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public virtual bool CanRedo
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "CanRedo");
			}
		}

		/// <summary>
		/// SupportByVersion MSForms 2
		/// Get
		/// </summary>
		[SupportByVersion("MSForms", 2)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public virtual bool CanUndo
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "CanUndo");
			}
		}

		/// <summary>
		/// SupportByVersion MSForms 2
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSForms", 2)]
		public virtual string Caption
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "Caption");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "Caption", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSForms 2
		/// Get
		/// </summary>
		[SupportByVersion("MSForms", 2)]
		public virtual NetOffice.MSFormsApi.Controls Controls
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.MSFormsApi.Controls>(this, "Controls", typeof(NetOffice.MSFormsApi.Controls));
			}
		}

		/// <summary>
		/// SupportByVersion MSForms 2
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSForms", 2)]
		public virtual NetOffice.MSFormsApi.Enums.fmCycle Cycle
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteEnumPropertyGet<NetOffice.MSFormsApi.Enums.fmCycle>(this, "Cycle");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteEnumPropertySet(this, "Cycle", value);
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
		/// Get
		/// </summary>
		[SupportByVersion("MSForms", 2)]
		public virtual Single InsideHeight
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteSinglePropertyGet(this, "InsideHeight");
			}
		}

		/// <summary>
		/// SupportByVersion MSForms 2
		/// Get
		/// </summary>
		[SupportByVersion("MSForms", 2)]
		public virtual Single InsideWidth
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteSinglePropertyGet(this, "InsideWidth");
			}
		}

		/// <summary>
		/// SupportByVersion MSForms 2
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSForms", 2)]
		public virtual NetOffice.MSFormsApi.Enums.fmScrollBars KeepScrollBarsVisible
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteEnumPropertyGet<NetOffice.MSFormsApi.Enums.fmScrollBars>(this, "KeepScrollBarsVisible");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteEnumPropertySet(this, "KeepScrollBarsVisible", value);
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
		public virtual NetOffice.MSFormsApi.Enums.fmPictureAlignment PictureAlignment
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteEnumPropertyGet<NetOffice.MSFormsApi.Enums.fmPictureAlignment>(this, "PictureAlignment");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteEnumPropertySet(this, "PictureAlignment", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSForms 2
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSForms", 2)]
		public virtual stdole.Picture Picture
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
		public virtual NetOffice.MSFormsApi.Enums.fmPictureSizeMode PictureSizeMode
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteEnumPropertyGet<NetOffice.MSFormsApi.Enums.fmPictureSizeMode>(this, "PictureSizeMode");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteEnumPropertySet(this, "PictureSizeMode", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSForms 2
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSForms", 2)]
		public virtual bool PictureTiling
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "PictureTiling");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "PictureTiling", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSForms 2
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSForms", 2)]
		public virtual NetOffice.MSFormsApi.Enums.fmScrollBars ScrollBars
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteEnumPropertyGet<NetOffice.MSFormsApi.Enums.fmScrollBars>(this, "ScrollBars");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteEnumPropertySet(this, "ScrollBars", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSForms 2
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSForms", 2)]
		public virtual Single ScrollHeight
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteSinglePropertyGet(this, "ScrollHeight");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "ScrollHeight", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSForms 2
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSForms", 2)]
		public virtual Single ScrollLeft
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteSinglePropertyGet(this, "ScrollLeft");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "ScrollLeft", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSForms 2
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSForms", 2)]
		public virtual Single ScrollTop
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteSinglePropertyGet(this, "ScrollTop");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "ScrollTop", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSForms 2
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSForms", 2)]
		public virtual Single ScrollWidth
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteSinglePropertyGet(this, "ScrollWidth");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "ScrollWidth", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSForms 2
		/// Get
		/// </summary>
		[SupportByVersion("MSForms", 2)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public virtual NetOffice.MSFormsApi.Controls Selected
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.MSFormsApi.Controls>(this, "Selected", typeof(NetOffice.MSFormsApi.Controls));
			}
		}

		/// <summary>
		/// SupportByVersion MSForms 2
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSForms", 2)]
		public virtual NetOffice.MSFormsApi.Enums.fmSpecialEffect SpecialEffect
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteEnumPropertyGet<NetOffice.MSFormsApi.Enums.fmSpecialEffect>(this, "SpecialEffect");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteEnumPropertySet(this, "SpecialEffect", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSForms 2
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSForms", 2)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public virtual NetOffice.MSFormsApi.Enums.fmVerticalScrollBarSide VerticalScrollBarSide
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteEnumPropertyGet<NetOffice.MSFormsApi.Enums.fmVerticalScrollBarSide>(this, "VerticalScrollBarSide");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteEnumPropertySet(this, "VerticalScrollBarSide", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSForms 2
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSForms", 2)]
		public virtual Int16 Zoom
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteInt16PropertyGet(this, "Zoom");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "Zoom", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSForms 2
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSForms", 2)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public virtual NetOffice.MSFormsApi.Enums.fmMode DesignMode
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteEnumPropertyGet<NetOffice.MSFormsApi.Enums.fmMode>(this, "DesignMode");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteEnumPropertySet(this, "DesignMode", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSForms 2
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSForms", 2)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public virtual NetOffice.MSFormsApi.Enums.fmMode ShowToolbox
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteEnumPropertyGet<NetOffice.MSFormsApi.Enums.fmMode>(this, "ShowToolbox");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteEnumPropertySet(this, "ShowToolbox", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSForms 2
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSForms", 2)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public virtual NetOffice.MSFormsApi.Enums.fmMode ShowGridDots
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteEnumPropertyGet<NetOffice.MSFormsApi.Enums.fmMode>(this, "ShowGridDots");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteEnumPropertySet(this, "ShowGridDots", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSForms 2
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSForms", 2)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public virtual NetOffice.MSFormsApi.Enums.fmMode SnapToGrid
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteEnumPropertyGet<NetOffice.MSFormsApi.Enums.fmMode>(this, "SnapToGrid");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteEnumPropertySet(this, "SnapToGrid", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSForms 2
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSForms", 2)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public virtual Single GridX
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteSinglePropertyGet(this, "GridX");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "GridX", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSForms 2
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSForms", 2)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public virtual Single GridY
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteSinglePropertyGet(this, "GridY");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "GridY", value);
			}
		}

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion MSForms 2
		/// </summary>
		/// <param name="insideHeight">Int32 insideHeight</param>
		[SupportByVersion("MSForms", 2)]
		public virtual void _GetInsideHeight(out Int32 insideHeight)
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
		public virtual void _GetInsideWidth(out Int32 insideWidth)
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
		public virtual void _SetScrollHeight(Int32 scrollHeight)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "_SetScrollHeight", scrollHeight);
		}

		/// <summary>
		/// SupportByVersion MSForms 2
		/// </summary>
		/// <param name="scrollHeight">Int32 scrollHeight</param>
		[SupportByVersion("MSForms", 2)]
		public virtual void _GetScrollHeight(out Int32 scrollHeight)
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
		public virtual void _SetScrollLeft(Int32 scrollLeft)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "_SetScrollLeft", scrollLeft);
		}

		/// <summary>
		/// SupportByVersion MSForms 2
		/// </summary>
		/// <param name="scrollLeft">Int32 scrollLeft</param>
		[SupportByVersion("MSForms", 2)]
		public virtual void _GetScrollLeft(out Int32 scrollLeft)
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
		public virtual void _SetScrollTop(Int32 scrollTop)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "_SetScrollTop", scrollTop);
		}

		/// <summary>
		/// SupportByVersion MSForms 2
		/// </summary>
		/// <param name="scrollTop">Int32 scrollTop</param>
		[SupportByVersion("MSForms", 2)]
		public virtual void _GetScrollTop(out Int32 scrollTop)
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
		public virtual void _SetScrollWidth(Int32 scrollWidth)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "_SetScrollWidth", scrollWidth);
		}

		/// <summary>
		/// SupportByVersion MSForms 2
		/// </summary>
		/// <param name="scrollWidth">Int32 scrollWidth</param>
		[SupportByVersion("MSForms", 2)]
		public virtual void _GetScrollWidth(out Int32 scrollWidth)
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
		public virtual void Copy()
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "Copy");
		}

		/// <summary>
		/// SupportByVersion MSForms 2
		/// </summary>
		[SupportByVersion("MSForms", 2)]
		public virtual void Cut()
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "Cut");
		}

		/// <summary>
		/// SupportByVersion MSForms 2
		/// </summary>
		[SupportByVersion("MSForms", 2)]
		public virtual void Paste()
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "Paste");
		}

		/// <summary>
		/// SupportByVersion MSForms 2
		/// </summary>
		[SupportByVersion("MSForms", 2)]
		public virtual void RedoAction()
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "RedoAction");
		}

		/// <summary>
		/// SupportByVersion MSForms 2
		/// </summary>
		[SupportByVersion("MSForms", 2)]
		public virtual void Repaint()
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "Repaint");
		}

		/// <summary>
		/// SupportByVersion MSForms 2
		/// </summary>
		/// <param name="xAction">optional object xAction</param>
		/// <param name="yAction">optional object yAction</param>
		[SupportByVersion("MSForms", 2)]
		public virtual void Scroll(object xAction, object yAction)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "Scroll", xAction, yAction);
		}

		/// <summary>
		/// SupportByVersion MSForms 2
		/// </summary>
		[CustomMethod]
		[SupportByVersion("MSForms", 2)]
		public virtual void Scroll()
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "Scroll");
		}

		/// <summary>
		/// SupportByVersion MSForms 2
		/// </summary>
		/// <param name="xAction">optional object xAction</param>
		[CustomMethod]
		[SupportByVersion("MSForms", 2)]
		public virtual void Scroll(object xAction)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "Scroll", xAction);
		}

		/// <summary>
		/// SupportByVersion MSForms 2
		/// </summary>
		[SupportByVersion("MSForms", 2)]
		public virtual void SetDefaultTabOrder()
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "SetDefaultTabOrder");
		}

		/// <summary>
		/// SupportByVersion MSForms 2
		/// </summary>
		[SupportByVersion("MSForms", 2)]
		public virtual void UndoAction()
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "UndoAction");
		}

		/// <summary>
		/// SupportByVersion MSForms 2
		/// </summary>
		/// <param name="gridX">Int32 gridX</param>
		[SupportByVersion("MSForms", 2)]
		public virtual void _SetGridX(Int32 gridX)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "_SetGridX", gridX);
		}

		/// <summary>
		/// SupportByVersion MSForms 2
		/// </summary>
		/// <param name="gridX">Int32 gridX</param>
		[SupportByVersion("MSForms", 2)]
		public virtual void _GetGridX(out Int32 gridX)
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
		public virtual void _SetGridY(Int32 gridY)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "_SetGridY", gridY);
		}

		/// <summary>
		/// SupportByVersion MSForms 2
		/// </summary>
		/// <param name="gridY">Int32 gridY</param>
		[SupportByVersion("MSForms", 2)]
		public virtual void _GetGridY(out Int32 gridY)
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


