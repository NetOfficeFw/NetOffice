using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.MSFormsApi.Behind
{
    /// <summary>
    /// IMdcCombo
    /// </summary>
    [SyntaxBypass]
    public class IMdcCombo_ : COMObject, NetOffice.MSFormsApi.IMdcCombo_
    {
        #region Ctor

        /// <summary>
        /// Stub Ctor, not intended to use
        /// </summary>
        public IMdcCombo_() : base()
        {
        }

        #endregion

        #region Properties

        /// <summary>
        /// SupportByVersion MSForms 2
        /// Get/Set
        /// </summary>
        /// <param name="pvargColumn">optional object pvargColumn</param>
        /// <param name="pvargIndex">optional object pvargIndex</param>
        [SupportByVersion("MSForms", 2)]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        public virtual object get_Column(object pvargColumn, object pvargIndex)
        {
            return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "Column", pvargColumn, pvargIndex);
        }

        /// <summary>
        /// SupportByVersion MSForms 2
        /// Get/Set
        /// </summary>
        /// <param name="pvargColumn">optional object pvargColumn</param>
        /// <param name="pvargIndex">optional object pvargIndex</param>
        /// <param name="value">optional object value</param>
        [SupportByVersion("MSForms", 2)]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        public virtual void set_Column(object pvargColumn, object pvargIndex, object value)
        {
            InvokerService.InvokeInternal.ExecutePropertySet(this, "Column", pvargColumn, pvargIndex, value);
        }

        /// <summary>
        /// SupportByVersion MSForms 2
        /// Alias for get_Column
        /// </summary>
        /// <param name="pvargColumn">optional object pvargColumn</param>
        /// <param name="pvargIndex">optional object pvargIndex</param>
        [SupportByVersion("MSForms", 2), Redirect("get_Column")]
        public virtual object Column(object pvargColumn, object pvargIndex)
        {
            return get_Column(pvargColumn, pvargIndex);
        }

        /// <summary>
        /// SupportByVersion MSForms 2
        /// Get/Set
        /// </summary>
        /// <param name="pvargColumn">optional object pvargColumn</param>
        [SupportByVersion("MSForms", 2)]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        public virtual object get_Column(object pvargColumn)
        {
            return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "Column", pvargColumn);
        }

        /// <summary>
        /// SupportByVersion MSForms 2
        /// Get/Set
        /// </summary>
        /// <param name="pvargColumn">optional object pvargColumn</param>
        /// <param name="value">optional object value</param>
        [SupportByVersion("MSForms", 2)]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        public virtual void set_Column(object pvargColumn, object value)
        {
            InvokerService.InvokeInternal.ExecutePropertySet(this, "Column", pvargColumn, value);
        }

        /// <summary>
        /// SupportByVersion MSForms 2
        /// Alias for get_Column
        /// </summary>
        /// <param name="pvargColumn">optional object pvargColumn</param>
        [SupportByVersion("MSForms", 2), Redirect("get_Column")]
        public virtual object Column(object pvargColumn)
        {
            return get_Column(pvargColumn);
        }

        /// <summary>
        /// SupportByVersion MSForms 2
        /// Get/Set
        /// </summary>
        /// <param name="pvargIndex">optional object pvargIndex</param>
        /// <param name="pvargColumn">optional object pvargColumn</param>
        [SupportByVersion("MSForms", 2)]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        public virtual object get_List(object pvargIndex, object pvargColumn)
        {
            return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "List", pvargIndex, pvargColumn);
        }

        /// <summary>
        /// SupportByVersion MSForms 2
        /// Get/Set
        /// </summary>
        /// <param name="pvargIndex">optional object pvargIndex</param>
        /// <param name="pvargColumn">optional object pvargColumn</param>
        /// <param name="value">optional object value</param>
        [SupportByVersion("MSForms", 2)]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        public virtual void set_List(object pvargIndex, object pvargColumn, object value)
        {
            InvokerService.InvokeInternal.ExecutePropertySet(this, "List", pvargIndex, pvargColumn, value);
        }

        /// <summary>
        /// SupportByVersion MSForms 2
        /// Alias for get_List
        /// </summary>
        /// <param name="pvargIndex">optional object pvargIndex</param>
        /// <param name="pvargColumn">optional object pvargColumn</param>
        [SupportByVersion("MSForms", 2), Redirect("get_List")]
        public virtual object List(object pvargIndex, object pvargColumn)
        {
            return get_List(pvargIndex, pvargColumn);
        }

        /// <summary>
        /// SupportByVersion MSForms 2
        /// Get/Set
        /// </summary>
        /// <param name="pvargIndex">optional object pvargIndex</param>
        [SupportByVersion("MSForms", 2)]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        public virtual object get_List(object pvargIndex)
        {
            return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "List", pvargIndex);
        }

        /// <summary>
        /// SupportByVersion MSForms 2
        /// Get/Set
        /// </summary>
        /// <param name="pvargIndex">optional object pvargIndex</param>
        /// <param name="value">optional object value</param>
        [SupportByVersion("MSForms", 2)]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        public virtual void set_List(object pvargIndex, object value)
        {
            InvokerService.InvokeInternal.ExecutePropertySet(this, "List", pvargIndex, value);
        }

        /// <summary>
        /// SupportByVersion MSForms 2
        /// Alias for get_List
        /// </summary>
        /// <param name="pvargIndex">optional object pvargIndex</param>
        [SupportByVersion("MSForms", 2), Redirect("get_List")]
        public virtual object List(object pvargIndex)
        {
            return get_List(pvargIndex);
        }

        #endregion
    }

    /// <summary>
    /// DispatchInterface IMdcCombo 
    /// SupportByVersion MSForms, 2
    /// </summary>
    [SupportByVersion("MSForms", 2)]
    [EntityType(EntityType.IsDispatchInterface), BaseType]
    public class IMdcCombo : IMdcCombo_, NetOffice.MSFormsApi.IMdcCombo
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
                    _contractType = typeof(NetOffice.MSFormsApi.IMdcCombo);
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
                    _type = typeof(IMdcCombo);
                return _type;
            }
        }

        #endregion

        #region Ctor

        /// <summary>
        /// Stub Ctor, not intended to use
        /// </summary>
        public IMdcCombo() : base()
        {
        }

        #endregion

        #region Properties

        /// <summary>
        /// SupportByVersion MSForms 2
        /// Get/Set
        /// </summary>
        [SupportByVersion("MSForms", 2)]
        public virtual bool AutoSize
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "AutoSize");
            }
            set
            {
                InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "AutoSize", value);
            }
        }

        /// <summary>
        /// SupportByVersion MSForms 2
        /// Get/Set
        /// </summary>
        [SupportByVersion("MSForms", 2)]
        public virtual bool AutoTab
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "AutoTab");
            }
            set
            {
                InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "AutoTab", value);
            }
        }

        /// <summary>
        /// SupportByVersion MSForms 2
        /// Get/Set
        /// </summary>
        [SupportByVersion("MSForms", 2)]
        public virtual bool AutoWordSelect
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "AutoWordSelect");
            }
            set
            {
                InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "AutoWordSelect", value);
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
        public virtual NetOffice.MSFormsApi.Enums.fmBackStyle BackStyle
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteEnumPropertyGet<NetOffice.MSFormsApi.Enums.fmBackStyle>(this, "BackStyle");
            }
            set
            {
                InvokerService.InvokeInternal.ExecuteEnumPropertySet(this, "BackStyle", value);
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
        /// Get/Set
        /// </summary>
        [SupportByVersion("MSForms", 2)]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        public virtual bool BordersSuppress
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "BordersSuppress");
            }
            set
            {
                InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "BordersSuppress", value);
            }
        }

        /// <summary>
        /// SupportByVersion MSForms 2
        /// Get/Set
        /// </summary>
        [SupportByVersion("MSForms", 2)]
        public virtual object BoundColumn
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "BoundColumn");
            }
            set
            {
                InvokerService.InvokeInternal.ExecuteVariantPropertySet(this, "BoundColumn", value);
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
        /// Get/Set
        /// </summary>
        [SupportByVersion("MSForms", 2)]
        public virtual Int32 ColumnCount
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteInt32PropertyGet(this, "ColumnCount");
            }
            set
            {
                InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "ColumnCount", value);
            }
        }

        /// <summary>
        /// SupportByVersion MSForms 2
        /// Get/Set
        /// </summary>
        [SupportByVersion("MSForms", 2)]
        public virtual bool ColumnHeads
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "ColumnHeads");
            }
            set
            {
                InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "ColumnHeads", value);
            }
        }

        /// <summary>
        /// SupportByVersion MSForms 2
        /// Get/Set
        /// </summary>
        [SupportByVersion("MSForms", 2)]
        public virtual string ColumnWidths
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "ColumnWidths");
            }
            set
            {
                InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "ColumnWidths", value);
            }
        }

        /// <summary>
        /// SupportByVersion MSForms 2
        /// Get
        /// </summary>
        [SupportByVersion("MSForms", 2)]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        public virtual Int32 CurTargetX
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteInt32PropertyGet(this, "CurTargetX");
            }
        }

        /// <summary>
        /// SupportByVersion MSForms 2
        /// Get
        /// </summary>
        [SupportByVersion("MSForms", 2)]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        public virtual Int32 CurTargetY
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteInt32PropertyGet(this, "CurTargetY");
            }
        }

        /// <summary>
        /// SupportByVersion MSForms 2
        /// Get/Set
        /// </summary>
        [SupportByVersion("MSForms", 2)]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        public virtual Int32 CurX
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteInt32PropertyGet(this, "CurX");
            }
            set
            {
                InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "CurX", value);
            }
        }

        /// <summary>
        /// SupportByVersion MSForms 2
        /// Get/Set
        /// </summary>
        [SupportByVersion("MSForms", 2)]
        public virtual NetOffice.MSFormsApi.Enums.fmDropButtonStyle DropButtonStyle
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteEnumPropertyGet<NetOffice.MSFormsApi.Enums.fmDropButtonStyle>(this, "DropButtonStyle");
            }
            set
            {
                InvokerService.InvokeInternal.ExecuteEnumPropertySet(this, "DropButtonStyle", value);
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
        public virtual bool HideSelection
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "HideSelection");
            }
            set
            {
                InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "HideSelection", value);
            }
        }

        /// <summary>
        /// SupportByVersion MSForms 2
        /// Get
        /// </summary>
        [SupportByVersion("MSForms", 2)]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        public virtual Int32 LineCount
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteInt32PropertyGet(this, "LineCount");
            }
        }

        /// <summary>
        /// SupportByVersion MSForms 2
        /// Get
        /// </summary>
        [SupportByVersion("MSForms", 2)]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        public virtual Int32 ListCount
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteInt32PropertyGet(this, "ListCount");
            }
        }

        /// <summary>
        /// SupportByVersion MSForms 2
        /// Get/Set
        /// Unknown COM Proxy
        /// </summary>
        [SupportByVersion("MSForms", 2), ProxyResult]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        public virtual object ListCursor
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteReferencePropertyGet(this, "ListCursor");
            }
            set
            {
                InvokerService.InvokeInternal.ExecuteReferencePropertySet(this, "ListCursor", value);
            }
        }

        /// <summary>
        /// SupportByVersion MSForms 2
        /// Get/Set
        /// </summary>
        [SupportByVersion("MSForms", 2)]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        public virtual object ListIndex
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "ListIndex");
            }
            set
            {
                InvokerService.InvokeInternal.ExecuteVariantPropertySet(this, "ListIndex", value);
            }
        }

        /// <summary>
        /// SupportByVersion MSForms 2
        /// Get/Set
        /// </summary>
        [SupportByVersion("MSForms", 2)]
        public virtual Int32 ListRows
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteInt32PropertyGet(this, "ListRows");
            }
            set
            {
                InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "ListRows", value);
            }
        }

        /// <summary>
        /// SupportByVersion MSForms 2
        /// Get/Set
        /// </summary>
        [SupportByVersion("MSForms", 2)]
        public virtual NetOffice.MSFormsApi.Enums.fmListStyle ListStyle
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteEnumPropertyGet<NetOffice.MSFormsApi.Enums.fmListStyle>(this, "ListStyle");
            }
            set
            {
                InvokerService.InvokeInternal.ExecuteEnumPropertySet(this, "ListStyle", value);
            }
        }

        /// <summary>
        /// SupportByVersion MSForms 2
        /// Get/Set
        /// </summary>
        [SupportByVersion("MSForms", 2)]
        public virtual object ListWidth
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "ListWidth");
            }
            set
            {
                InvokerService.InvokeInternal.ExecuteVariantPropertySet(this, "ListWidth", value);
            }
        }

        /// <summary>
        /// SupportByVersion MSForms 2
        /// Get/Set
        /// </summary>
        [SupportByVersion("MSForms", 2)]
        public virtual bool Locked
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "Locked");
            }
            set
            {
                InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "Locked", value);
            }
        }

        /// <summary>
        /// SupportByVersion MSForms 2
        /// Get/Set
        /// </summary>
        [SupportByVersion("MSForms", 2)]
        public virtual NetOffice.MSFormsApi.Enums.fmMatchEntry MatchEntry
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteEnumPropertyGet<NetOffice.MSFormsApi.Enums.fmMatchEntry>(this, "MatchEntry");
            }
            set
            {
                InvokerService.InvokeInternal.ExecuteEnumPropertySet(this, "MatchEntry", value);
            }
        }

        /// <summary>
        /// SupportByVersion MSForms 2
        /// Get
        /// </summary>
        [SupportByVersion("MSForms", 2)]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        public virtual bool MatchFound
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "MatchFound");
            }
        }

        /// <summary>
        /// SupportByVersion MSForms 2
        /// Get/Set
        /// </summary>
        [SupportByVersion("MSForms", 2)]
        public virtual bool MatchRequired
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "MatchRequired");
            }
            set
            {
                InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "MatchRequired", value);
            }
        }

        /// <summary>
        /// SupportByVersion MSForms 2
        /// Get/Set
        /// </summary>
        [SupportByVersion("MSForms", 2)]
        public virtual Int32 MaxLength
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteInt32PropertyGet(this, "MaxLength");
            }
            set
            {
                InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "MaxLength", value);
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
        public virtual bool SelectionMargin
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "SelectionMargin");
            }
            set
            {
                InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "SelectionMargin", value);
            }
        }

        /// <summary>
        /// SupportByVersion MSForms 2
        /// Get/Set
        /// </summary>
        [SupportByVersion("MSForms", 2)]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        public virtual Int32 SelLength
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteInt32PropertyGet(this, "SelLength");
            }
            set
            {
                InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "SelLength", value);
            }
        }

        /// <summary>
        /// SupportByVersion MSForms 2
        /// Get/Set
        /// </summary>
        [SupportByVersion("MSForms", 2)]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        public virtual Int32 SelStart
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteInt32PropertyGet(this, "SelStart");
            }
            set
            {
                InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "SelStart", value);
            }
        }

        /// <summary>
        /// SupportByVersion MSForms 2
        /// Get/Set
        /// </summary>
        [SupportByVersion("MSForms", 2)]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        public virtual string SelText
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "SelText");
            }
            set
            {
                InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "SelText", value);
            }
        }

        /// <summary>
        /// SupportByVersion MSForms 2
        /// Get/Set
        /// </summary>
        [SupportByVersion("MSForms", 2)]
        public virtual NetOffice.MSFormsApi.Enums.fmShowDropButtonWhen ShowDropButtonWhen
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteEnumPropertyGet<NetOffice.MSFormsApi.Enums.fmShowDropButtonWhen>(this, "ShowDropButtonWhen");
            }
            set
            {
                InvokerService.InvokeInternal.ExecuteEnumPropertySet(this, "ShowDropButtonWhen", value);
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
        public virtual NetOffice.MSFormsApi.Enums.fmStyle Style
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteEnumPropertyGet<NetOffice.MSFormsApi.Enums.fmStyle>(this, "Style");
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
        public virtual string Text
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "Text");
            }
            set
            {
                InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "Text", value);
            }
        }

        /// <summary>
        /// SupportByVersion MSForms 2
        /// Get/Set
        /// </summary>
        [SupportByVersion("MSForms", 2)]
        public virtual NetOffice.MSFormsApi.Enums.fmTextAlign TextAlign
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteEnumPropertyGet<NetOffice.MSFormsApi.Enums.fmTextAlign>(this, "TextAlign");
            }
            set
            {
                InvokerService.InvokeInternal.ExecuteEnumPropertySet(this, "TextAlign", value);
            }
        }

        /// <summary>
        /// SupportByVersion MSForms 2
        /// Get/Set
        /// </summary>
        [SupportByVersion("MSForms", 2)]
        public virtual object TextColumn
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "TextColumn");
            }
            set
            {
                InvokerService.InvokeInternal.ExecuteVariantPropertySet(this, "TextColumn", value);
            }
        }

        /// <summary>
        /// SupportByVersion MSForms 2
        /// Get
        /// </summary>
        [SupportByVersion("MSForms", 2)]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        public virtual Int32 TextLength
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteInt32PropertyGet(this, "TextLength");
            }
        }

        /// <summary>
        /// SupportByVersion MSForms 2
        /// Get/Set
        /// </summary>
        [SupportByVersion("MSForms", 2)]
        public virtual object TopIndex
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "TopIndex");
            }
            set
            {
                InvokerService.InvokeInternal.ExecuteVariantPropertySet(this, "TopIndex", value);
            }
        }

        /// <summary>
        /// SupportByVersion MSForms 2
        /// Get
        /// </summary>
        [SupportByVersion("MSForms", 2)]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        public virtual bool Valid
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "Valid");
            }
        }

        /// <summary>
        /// SupportByVersion MSForms 2
        /// Get/Set
        /// </summary>
        [SupportByVersion("MSForms", 2)]
        public virtual object Value
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "Value");
            }
            set
            {
                InvokerService.InvokeInternal.ExecuteVariantPropertySet(this, "Value", value);
            }
        }

        /// <summary>
        /// SupportByVersion MSForms 2
        /// Get/Set
        /// </summary>
        [SupportByVersion("MSForms", 2)]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        public virtual object Column
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "Column");
            }
            set
            {
                InvokerService.InvokeInternal.ExecuteVariantPropertySet(this, "Column", value);
            }
        }

        /// <summary>
        /// SupportByVersion MSForms 2
        /// Get/Set
        /// </summary>
        [SupportByVersion("MSForms", 2)]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        public virtual object List
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "List");
            }
            set
            {
                InvokerService.InvokeInternal.ExecuteVariantPropertySet(this, "List", value);
            }
        }

        /// <summary>
        /// SupportByVersion MSForms 2
        /// Get/Set
        /// </summary>
        [SupportByVersion("MSForms", 2)]
        public virtual NetOffice.MSFormsApi.Enums.fmIMEMode IMEMode
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteEnumPropertyGet<NetOffice.MSFormsApi.Enums.fmIMEMode>(this, "IMEMode");
            }
            set
            {
                InvokerService.InvokeInternal.ExecuteEnumPropertySet(this, "IMEMode", value);
            }
        }

        /// <summary>
        /// SupportByVersion MSForms 2
        /// Get/Set
        /// </summary>
        [SupportByVersion("MSForms", 2)]
        public virtual NetOffice.MSFormsApi.Enums.fmEnterFieldBehavior EnterFieldBehavior
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteEnumPropertyGet<NetOffice.MSFormsApi.Enums.fmEnterFieldBehavior>(this, "EnterFieldBehavior");
            }
            set
            {
                InvokerService.InvokeInternal.ExecuteEnumPropertySet(this, "EnterFieldBehavior", value);
            }
        }

        /// <summary>
        /// SupportByVersion MSForms 2
        /// Get/Set
        /// </summary>
        [SupportByVersion("MSForms", 2)]
        public virtual NetOffice.MSFormsApi.Enums.fmDragBehavior DragBehavior
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteEnumPropertyGet<NetOffice.MSFormsApi.Enums.fmDragBehavior>(this, "DragBehavior");
            }
            set
            {
                InvokerService.InvokeInternal.ExecuteEnumPropertySet(this, "DragBehavior", value);
            }
        }

        /// <summary>
        /// SupportByVersion MSForms 2
        /// Get
        /// </summary>
        [SupportByVersion("MSForms", 2)]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        public virtual NetOffice.MSFormsApi.Enums.fmDisplayStyle DisplayStyle
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteEnumPropertyGet<NetOffice.MSFormsApi.Enums.fmDisplayStyle>(this, "DisplayStyle");
            }
        }

        #endregion

        #region Methods

        /// <summary>
        /// SupportByVersion MSForms 2
        /// </summary>
        /// <param name="pvargItem">optional object pvargItem</param>
        /// <param name="pvargIndex">optional object pvargIndex</param>
        [SupportByVersion("MSForms", 2)]
        public virtual void AddItem(object pvargItem, object pvargIndex)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "AddItem", pvargItem, pvargIndex);
        }

        /// <summary>
        /// SupportByVersion MSForms 2
        /// </summary>
        [CustomMethod]
        [SupportByVersion("MSForms", 2)]
        public virtual void AddItem()
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "AddItem");
        }

        /// <summary>
        /// SupportByVersion MSForms 2
        /// </summary>
        /// <param name="pvargItem">optional object pvargItem</param>
        [CustomMethod]
        [SupportByVersion("MSForms", 2)]
        public virtual void AddItem(object pvargItem)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "AddItem", pvargItem);
        }

        /// <summary>
        /// SupportByVersion MSForms 2
        /// </summary>
        [SupportByVersion("MSForms", 2)]
        public virtual void Clear()
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "Clear");
        }

        /// <summary>
        /// SupportByVersion MSForms 2
        /// </summary>
        [SupportByVersion("MSForms", 2)]
        public virtual void DropDown()
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "DropDown");
        }

        /// <summary>
        /// SupportByVersion MSForms 2
        /// </summary>
        /// <param name="pvargIndex">object pvargIndex</param>
        [SupportByVersion("MSForms", 2)]
        public virtual void RemoveItem(object pvargIndex)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "RemoveItem", pvargIndex);
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

        #endregion

        #pragma warning restore
    }
}
