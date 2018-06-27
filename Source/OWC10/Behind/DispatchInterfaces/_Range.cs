using System.Collections;
using System.Collections.Generic;
using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;
using NetOffice.CollectionsGeneric;

namespace NetOffice.OWC10Api.Behind
{
    /// <summary>
    /// _Range
    /// </summary>
    [SyntaxBypass]
    public class _Range_ : COMObject, NetOffice.OWC10Api._Range_
    {
        #region Ctor

        /// <summary>
        /// Stub Ctor, not intended to use
        /// </summary>
        public _Range_() : base()
        {
        }

        #endregion

        #region Properties

        /// <summary>
        /// SupportByVersion OWC10 1
        /// Get
        /// </summary>
        /// <param name="rowAbsolute">optional object rowAbsolute</param>
        /// <param name="columnAbsolute">optional object columnAbsolute</param>
        /// <param name="referenceStyle">optional NetOffice.OWC10Api.Enums.XlReferenceStyle referenceStyle</param>
        /// <param name="external">optional object external</param>
        /// <param name="relativeTo">optional object relativeTo</param>
        [SupportByVersion("OWC10", 1)]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        public virtual string get_Address(object rowAbsolute, object columnAbsolute, object referenceStyle, object external, object relativeTo)
        {
            return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "Address", new object[] { rowAbsolute, columnAbsolute, referenceStyle, external, relativeTo });
        }

        /// <summary>
        /// SupportByVersion OWC10 1
        /// Alias for get_Address
        /// </summary>
        /// <param name="rowAbsolute">optional object rowAbsolute</param>
        /// <param name="columnAbsolute">optional object columnAbsolute</param>
        /// <param name="referenceStyle">optional NetOffice.OWC10Api.Enums.XlReferenceStyle referenceStyle</param>
        /// <param name="external">optional object external</param>
        /// <param name="relativeTo">optional object relativeTo</param>
        [SupportByVersion("OWC10", 1), Redirect("get_Address")]
        public virtual string Address(object rowAbsolute, object columnAbsolute, object referenceStyle, object external, object relativeTo)
        {
            return get_Address(rowAbsolute, columnAbsolute, referenceStyle, external, relativeTo);
        }

        /// <summary>
        /// SupportByVersion OWC10 1
        /// Get
        /// </summary>
        /// <param name="rowAbsolute">optional object rowAbsolute</param>
        [SupportByVersion("OWC10", 1)]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        public virtual string get_Address(object rowAbsolute)
        {
            return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "Address", rowAbsolute);
        }

        /// <summary>
        /// SupportByVersion OWC10 1
        /// Alias for get_Address
        /// </summary>
        /// <param name="rowAbsolute">optional object rowAbsolute</param>
        [SupportByVersion("OWC10", 1), Redirect("get_Address")]
        public virtual string Address(object rowAbsolute)
        {
            return get_Address(rowAbsolute);
        }

        /// <summary>
        /// SupportByVersion OWC10 1
        /// Get
        /// </summary>
        /// <param name="rowAbsolute">optional object rowAbsolute</param>
        /// <param name="columnAbsolute">optional object columnAbsolute</param>
        [SupportByVersion("OWC10", 1)]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        public virtual string get_Address(object rowAbsolute, object columnAbsolute)
        {
            return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "Address", rowAbsolute, columnAbsolute);
        }

        /// <summary>
        /// SupportByVersion OWC10 1
        /// Alias for get_Address
        /// </summary>
        /// <param name="rowAbsolute">optional object rowAbsolute</param>
        /// <param name="columnAbsolute">optional object columnAbsolute</param>
        [SupportByVersion("OWC10", 1), Redirect("get_Address")]
        public virtual string Address(object rowAbsolute, object columnAbsolute)
        {
            return get_Address(rowAbsolute, columnAbsolute);
        }

        /// <summary>
        /// SupportByVersion OWC10 1
        /// Get
        /// </summary>
        /// <param name="rowAbsolute">optional object rowAbsolute</param>
        /// <param name="columnAbsolute">optional object columnAbsolute</param>
        /// <param name="referenceStyle">optional NetOffice.OWC10Api.Enums.XlReferenceStyle referenceStyle</param>
        [SupportByVersion("OWC10", 1)]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        public virtual string get_Address(object rowAbsolute, object columnAbsolute, object referenceStyle)
        {
            return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "Address", rowAbsolute, columnAbsolute, referenceStyle);
        }

        /// <summary>
        /// SupportByVersion OWC10 1
        /// Alias for get_Address
        /// </summary>
        /// <param name="rowAbsolute">optional object rowAbsolute</param>
        /// <param name="columnAbsolute">optional object columnAbsolute</param>
        /// <param name="referenceStyle">optional NetOffice.OWC10Api.Enums.XlReferenceStyle referenceStyle</param>
        [SupportByVersion("OWC10", 1), Redirect("get_Address")]
        public virtual string Address(object rowAbsolute, object columnAbsolute, object referenceStyle)
        {
            return get_Address(rowAbsolute, columnAbsolute, referenceStyle);
        }

        /// <summary>
        /// SupportByVersion OWC10 1
        /// Get
        /// </summary>
        /// <param name="rowAbsolute">optional object rowAbsolute</param>
        /// <param name="columnAbsolute">optional object columnAbsolute</param>
        /// <param name="referenceStyle">optional NetOffice.OWC10Api.Enums.XlReferenceStyle referenceStyle</param>
        /// <param name="external">optional object external</param>
        [SupportByVersion("OWC10", 1)]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        public virtual string get_Address(object rowAbsolute, object columnAbsolute, object referenceStyle, object external)
        {
            return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "Address", rowAbsolute, columnAbsolute, referenceStyle, external);
        }

        /// <summary>
        /// SupportByVersion OWC10 1
        /// Alias for get_Address
        /// </summary>
        /// <param name="rowAbsolute">optional object rowAbsolute</param>
        /// <param name="columnAbsolute">optional object columnAbsolute</param>
        /// <param name="referenceStyle">optional NetOffice.OWC10Api.Enums.XlReferenceStyle referenceStyle</param>
        /// <param name="external">optional object external</param>
        [SupportByVersion("OWC10", 1), Redirect("get_Address")]
        public virtual string Address(object rowAbsolute, object columnAbsolute, object referenceStyle, object external)
        {
            return get_Address(rowAbsolute, columnAbsolute, referenceStyle, external);
        }

        /// <summary>
        /// SupportByVersion OWC10 1
        /// Get
        /// </summary>
        /// <param name="rowOffset">optional object rowOffset</param>
        /// <param name="columnOffset">optional object columnOffset</param>
        [SupportByVersion("OWC10", 1)]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        public virtual NetOffice.OWC10Api._Range get_Offset(object rowOffset, object columnOffset)
        {
            return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.OWC10Api._Range>(this, "Offset", typeof(NetOffice.OWC10Api._Range), rowOffset, columnOffset);
        }

        /// <summary>
        /// SupportByVersion OWC10 1
        /// Alias for get_Offset
        /// </summary>
        /// <param name="rowOffset">optional object rowOffset</param>
        /// <param name="columnOffset">optional object columnOffset</param>
        [SupportByVersion("OWC10", 1), Redirect("get_Offset")]
        public virtual NetOffice.OWC10Api._Range Offset(object rowOffset, object columnOffset)
        {
            return get_Offset(rowOffset, columnOffset);
        }

        /// <summary>
        /// SupportByVersion OWC10 1
        /// Get
        /// </summary>
        /// <param name="rowOffset">optional object rowOffset</param>
        [SupportByVersion("OWC10", 1)]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        public virtual NetOffice.OWC10Api._Range get_Offset(object rowOffset)
        {
            return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.OWC10Api._Range>(this, "Offset", typeof(NetOffice.OWC10Api._Range), rowOffset);
        }

        /// <summary>
        /// SupportByVersion OWC10 1
        /// Alias for get_Offset
        /// </summary>
        /// <param name="rowOffset">optional object rowOffset</param>
        [SupportByVersion("OWC10", 1), Redirect("get_Offset")]
        public virtual NetOffice.OWC10Api._Range Offset(object rowOffset)
        {
            return get_Offset(rowOffset);
        }

        /// <summary>
        /// SupportByVersion OWC10 1
        /// Get/Set
        /// </summary>
        /// <param name="rangeValueDataType">optional object rangeValueDataType</param>
        [SupportByVersion("OWC10", 1)]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        public virtual object get_Value(object rangeValueDataType)
        {
            return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "Value", rangeValueDataType);
        }

        /// <summary>
        /// SupportByVersion OWC10 1
        /// Get/Set
        /// </summary>
        /// <param name="rangeValueDataType">optional object rangeValueDataType</param>
        /// <param name="value">optional object value</param>
        [SupportByVersion("OWC10", 1)]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        public virtual void set_Value(object rangeValueDataType, object value)
        {
            InvokerService.InvokeInternal.ExecutePropertySet(this, "Value", rangeValueDataType, value);
        }

        /// <summary>
        /// SupportByVersion OWC10 1
        /// Alias for get_Value
        /// </summary>
        /// <param name="rangeValueDataType">optional object rangeValueDataType</param>
        [SupportByVersion("OWC10", 1), Redirect("get_Value")]
        public virtual object Value(object rangeValueDataType)
        {
            return get_Value(rangeValueDataType);
        }

        #endregion
    }

    /// <summary>
    /// DispatchInterface _Range 
    /// SupportByVersion OWC10, 1
    /// </summary>
    public class _Range : _Range_, NetOffice.OWC10Api._Range
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
                    _contractType = typeof(NetOffice.OWC10Api._Range);
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
                    _type = typeof(_Range);
                return _type;
            }
        }

        #endregion

        #region Ctor

        /// <summary>
        /// Stub Ctor, not intended to use
        /// </summary>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        public _Range() : base()
        {
        }

        #endregion

        #region Properties

        /// <summary>
        /// SupportByVersion OWC10 1
        /// Get/Set
        /// Custom Indexer
        /// </summary>
        /// <param name="row">optional object row</param>
        [SupportByVersion("OWC10", 1)]
        [NetRuntimeSystem.Runtime.CompilerServices.IndexerName("Item"), IndexProperty, CustomIndexer]
        public virtual object this[object row]
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "_Default", row);
            }
            set
            {
                InvokerService.InvokeInternal.ExecuteVariantPropertySet(this, "_Default", value, row);
            }
        }

        /// <summary>
        /// SupportByVersion OWC10 1
        /// Get/Set
        /// </summary>
        /// <param name="row">optional object row</param>
        /// <param name="column">optional object column</param>
        [SupportByVersion("OWC10", 1)]
        [NetRuntimeSystem.Runtime.CompilerServices.IndexerName("Item"), IndexProperty]
        public virtual object this[object row, object column]
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "_Default", row, column);
            }
            set
            {
                InvokerService.InvokeInternal.ExecuteVariantPropertySet(this, "_Default", value, row, column);
            }
        }

        /// <summary>
        /// SupportByVersion OWC10 1
        /// Get
        /// </summary>
        [SupportByVersion("OWC10", 1)]
        public virtual string Address
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "Address");
            }
        }

        /// <summary>
        /// SupportByVersion OWC10 1
        /// Get
        /// </summary>
        [SupportByVersion("OWC10", 1)]
        [BaseResult]
        public virtual NetOffice.OWC10Api.ISpreadsheet Application
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteBaseReferencePropertyGet<NetOffice.OWC10Api.ISpreadsheet>(this, "Application");
            }
        }

        /// <summary>
        /// SupportByVersion OWC10 1
        /// Get
        /// </summary>
        [SupportByVersion("OWC10", 1)]
        public virtual NetOffice.OWC10Api.Borders Borders
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.OWC10Api.Borders>(this, "Borders", typeof(NetOffice.OWC10Api.Borders));
            }
        }

        /// <summary>
        /// SupportByVersion OWC10 1
        /// Get
        /// </summary>
        [SupportByVersion("OWC10", 1)]
        [BaseResult]
        public virtual NetOffice.OWC10Api._Range Cells
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteBaseReferencePropertyGet<NetOffice.OWC10Api._Range>(this, "Cells");
            }
        }

        /// <summary>
        /// SupportByVersion OWC10 1
        /// Get
        /// </summary>
        [SupportByVersion("OWC10", 1)]
        public virtual Int32 Column
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteInt32PropertyGet(this, "Column");
            }
        }

        /// <summary>
        /// SupportByVersion OWC10 1
        /// Get
        /// </summary>
        [SupportByVersion("OWC10", 1)]
        [BaseResult]
        public virtual NetOffice.OWC10Api._Range Columns
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteBaseReferencePropertyGet<NetOffice.OWC10Api._Range>(this, "Columns");
            }
        }

        /// <summary>
        /// SupportByVersion OWC10 1
        /// Get/Set
        /// </summary>
        [SupportByVersion("OWC10", 1)]
        public virtual object ColumnWidth
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "ColumnWidth");
            }
            set
            {
                InvokerService.InvokeInternal.ExecuteVariantPropertySet(this, "ColumnWidth", value);
            }
        }

        /// <summary>
        /// SupportByVersion OWC10 1
        /// Get
        /// </summary>
        [SupportByVersion("OWC10", 1)]
        public virtual Int32 Count
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteInt32PropertyGet(this, "Count");
            }
        }

        /// <summary>
        /// SupportByVersion OWC10 1
        /// Get
        /// </summary>
        [SupportByVersion("OWC10", 1)]
        [BaseResult]
        public virtual NetOffice.OWC10Api._Range CurrentArray
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteBaseReferencePropertyGet<NetOffice.OWC10Api._Range>(this, "CurrentArray");
            }
        }

        /// <summary>
        /// SupportByVersion OWC10 1
        /// Get
        /// </summary>
        [SupportByVersion("OWC10", 1)]
        [BaseResult]
        public virtual NetOffice.OWC10Api._Range CurrentRegion
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteBaseReferencePropertyGet<NetOffice.OWC10Api._Range>(this, "CurrentRegion");
            }
        }

        /// <summary>
        /// SupportByVersion OWC10 1
        /// Get
        /// </summary>
        /// <param name="direction">NetOffice.OWC10Api.Enums.XlDirection direction</param>
        [SupportByVersion("OWC10", 1)]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        public virtual NetOffice.OWC10Api._Range get_End(NetOffice.OWC10Api.Enums.XlDirection direction)
        {
            return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.OWC10Api._Range>(this, "End", typeof(NetOffice.OWC10Api._Range), direction);
        }

        /// <summary>
        /// SupportByVersion OWC10 1
        /// Alias for get_End
        /// </summary>
        /// <param name="direction">NetOffice.OWC10Api.Enums.XlDirection direction</param>
        [SupportByVersion("OWC10", 1), Redirect("get_End")]
        public virtual NetOffice.OWC10Api._Range End(NetOffice.OWC10Api.Enums.XlDirection direction)
        {
            return get_End(direction);
        }

        /// <summary>
        /// SupportByVersion OWC10 1
        /// Get
        /// </summary>
        [SupportByVersion("OWC10", 1)]
        [BaseResult]
        public virtual NetOffice.OWC10Api._Range EntireColumn
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteBaseReferencePropertyGet<NetOffice.OWC10Api._Range>(this, "EntireColumn");
            }
        }

        /// <summary>
        /// SupportByVersion OWC10 1
        /// Get
        /// </summary>
        [SupportByVersion("OWC10", 1)]
        [BaseResult]
        public virtual NetOffice.OWC10Api._Range EntireRow
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteBaseReferencePropertyGet<NetOffice.OWC10Api._Range>(this, "EntireRow");
            }
        }

        /// <summary>
        /// SupportByVersion OWC10 1
        /// Get
        /// </summary>
        [SupportByVersion("OWC10", 1)]
        public virtual NetOffice.OWC10Api.Font Font
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.OWC10Api.Font>(this, "Font", typeof(NetOffice.OWC10Api.Font));
            }
        }

        /// <summary>
        /// SupportByVersion OWC10 1
        /// Get/Set
        /// </summary>
        [SupportByVersion("OWC10", 1)]
        public virtual object Formula
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "Formula");
            }
            set
            {
                InvokerService.InvokeInternal.ExecuteVariantPropertySet(this, "Formula", value);
            }
        }

        /// <summary>
        /// SupportByVersion OWC10 1
        /// Get/Set
        /// </summary>
        [SupportByVersion("OWC10", 1)]
        public virtual object FormulaArray
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "FormulaArray");
            }
            set
            {
                InvokerService.InvokeInternal.ExecuteVariantPropertySet(this, "FormulaArray", value);
            }
        }

        /// <summary>
        /// SupportByVersion OWC10 1
        /// Get/Set
        /// </summary>
        [SupportByVersion("OWC10", 1)]
        public virtual object FormulaLocal
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "FormulaLocal");
            }
            set
            {
                InvokerService.InvokeInternal.ExecuteVariantPropertySet(this, "FormulaLocal", value);
            }
        }

        /// <summary>
        /// SupportByVersion OWC10 1
        /// Get
        /// </summary>
        [SupportByVersion("OWC10", 1)]
        public virtual object HasArray
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "HasArray");
            }
        }

        /// <summary>
        /// SupportByVersion OWC10 1
        /// Get
        /// </summary>
        [SupportByVersion("OWC10", 1)]
        public virtual object HasFormula
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "HasFormula");
            }
        }

        /// <summary>
        /// SupportByVersion OWC10 1
        /// Get
        /// </summary>
        [SupportByVersion("OWC10", 1)]
        public virtual object Height
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "Height");
            }
        }

        /// <summary>
        /// SupportByVersion OWC10 1
        /// Get/Set
        /// </summary>
        [SupportByVersion("OWC10", 1)]
        public virtual bool Hidden
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "Hidden");
            }
            set
            {
                InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "Hidden", value);
            }
        }

        /// <summary>
        /// SupportByVersion OWC10 1
        /// Get/Set
        /// </summary>
        [SupportByVersion("OWC10", 1)]
        public virtual object HorizontalAlignment
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "HorizontalAlignment");
            }
            set
            {
                InvokerService.InvokeInternal.ExecuteVariantPropertySet(this, "HorizontalAlignment", value);
            }
        }

        /// <summary>
        /// SupportByVersion OWC10 1
        /// Get
        /// </summary>
        [SupportByVersion("OWC10", 1)]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        public virtual string HTMLData
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "HTMLData");
            }
        }

        /// <summary>
        /// SupportByVersion OWC10 1
        /// Get
        /// </summary>
        [SupportByVersion("OWC10", 1)]
        public virtual NetOffice.OWC10Api.Hyperlink Hyperlink
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.OWC10Api.Hyperlink>(this, "Hyperlink", typeof(NetOffice.OWC10Api.Hyperlink));
            }
        }

        /// <summary>
        /// SupportByVersion OWC10 1
        /// Get
        /// </summary>
        [SupportByVersion("OWC10", 1)]
        public virtual NetOffice.OWC10Api.Interior Interior
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.OWC10Api.Interior>(this, "Interior", typeof(NetOffice.OWC10Api.Interior));
            }
        }

        /// <summary>
        /// SupportByVersion OWC10 1
        /// Get
        /// </summary>
        [SupportByVersion("OWC10", 1)]
        public virtual object Left
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "Left");
            }
        }

        /// <summary>
        /// SupportByVersion OWC10 1
        /// Get/Set
        /// </summary>
        [SupportByVersion("OWC10", 1)]
        public virtual object Locked
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "Locked");
            }
            set
            {
                InvokerService.InvokeInternal.ExecuteVariantPropertySet(this, "Locked", value);
            }
        }

        /// <summary>
        /// SupportByVersion OWC10 1
        /// Get
        /// </summary>
        [SupportByVersion("OWC10", 1)]
        [BaseResult]
        public virtual NetOffice.OWC10Api._Range MergeArea
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteBaseReferencePropertyGet<NetOffice.OWC10Api._Range>(this, "MergeArea");
            }
        }

        /// <summary>
        /// SupportByVersion OWC10 1
        /// Get/Set
        /// </summary>
        [SupportByVersion("OWC10", 1)]
        public virtual object MergeCells
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "MergeCells");
            }
            set
            {
                InvokerService.InvokeInternal.ExecuteVariantPropertySet(this, "MergeCells", value);
            }
        }

        /// <summary>
        /// SupportByVersion OWC10 1
        /// Get/Set
        /// </summary>
        [SupportByVersion("OWC10", 1)]
        public virtual NetOffice.OWC10Api.Name Name
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.OWC10Api.Name>(this, "Name", typeof(NetOffice.OWC10Api.Name));
            }
            set
            {
                InvokerService.InvokeInternal.ExecuteReferencePropertySet(this, "Name", value);
            }
        }

        /// <summary>
        /// SupportByVersion OWC10 1
        /// Get
        /// </summary>
        [SupportByVersion("OWC10", 1)]
        [BaseResult]
        public virtual NetOffice.OWC10Api._Range Next
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteBaseReferencePropertyGet<NetOffice.OWC10Api._Range>(this, "Next");
            }
        }

        /// <summary>
        /// SupportByVersion OWC10 1
        /// Get/Set
        /// </summary>
        [SupportByVersion("OWC10", 1)]
        public virtual object NumberFormat
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "NumberFormat");
            }
            set
            {
                InvokerService.InvokeInternal.ExecuteVariantPropertySet(this, "NumberFormat", value);
            }
        }

        /// <summary>
        /// SupportByVersion OWC10 1
        /// Get
        /// </summary>
        [SupportByVersion("OWC10", 1)]
        [BaseResult]
        public virtual NetOffice.OWC10Api._Range Offset
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteBaseReferencePropertyGet<NetOffice.OWC10Api._Range>(this, "Offset");
            }
        }

        /// <summary>
        /// SupportByVersion OWC10 1
        /// Get
        /// </summary>
        [SupportByVersion("OWC10", 1)]
        public virtual NetOffice.OWC10Api.Worksheet Parent
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.OWC10Api.Worksheet>(this, "Parent", typeof(NetOffice.OWC10Api.Worksheet));
            }
        }

        /// <summary>
        /// SupportByVersion OWC10 1
        /// Get
        /// </summary>
        [SupportByVersion("OWC10", 1)]
        public virtual object PrefixCharacter
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "PrefixCharacter");
            }
        }

        /// <summary>
        /// SupportByVersion OWC10 1
        /// Get
        /// </summary>
        [SupportByVersion("OWC10", 1)]
        [BaseResult]
        public virtual NetOffice.OWC10Api._Range Previous
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteBaseReferencePropertyGet<NetOffice.OWC10Api._Range>(this, "Previous");
            }
        }

        /// <summary>
        /// SupportByVersion OWC10 1
        /// Get
        /// </summary>
        /// <param name="cell1">object cell1</param>
        /// <param name="cell2">optional object cell2</param>
        [SupportByVersion("OWC10", 1)]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        public virtual NetOffice.OWC10Api._Range get_Range(object cell1, object cell2)
        {
            return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.OWC10Api._Range>(this, "Range", typeof(NetOffice.OWC10Api._Range), cell1, cell2);
        }

        /// <summary>
        /// SupportByVersion OWC10 1
        /// Alias for get_Range
        /// </summary>
        /// <param name="cell1">object cell1</param>
        /// <param name="cell2">optional object cell2</param>
        [SupportByVersion("OWC10", 1), Redirect("get_Range")]
        public virtual NetOffice.OWC10Api._Range Range(object cell1, object cell2)
        {
            return get_Range(cell1, cell2);
        }

        /// <summary>
        /// SupportByVersion OWC10 1
        /// Get
        /// </summary>
        /// <param name="cell1">object cell1</param>
        [SupportByVersion("OWC10", 1)]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        public virtual NetOffice.OWC10Api._Range get_Range(object cell1)
        {
            return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.OWC10Api._Range>(this, "Range", typeof(NetOffice.OWC10Api._Range), cell1);
        }

        /// <summary>
        /// SupportByVersion OWC10 1
        /// Alias for get_Range
        /// </summary>
        /// <param name="cell1">object cell1</param>
        [SupportByVersion("OWC10", 1), Redirect("get_Range")]
        public virtual NetOffice.OWC10Api._Range Range(object cell1)
        {
            return get_Range(cell1);
        }

        /// <summary>
        /// SupportByVersion OWC10 1
        /// Get/Set
        /// </summary>
        [SupportByVersion("OWC10", 1)]
        public virtual object ReadingOrder
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "ReadingOrder");
            }
            set
            {
                InvokerService.InvokeInternal.ExecuteVariantPropertySet(this, "ReadingOrder", value);
            }
        }

        /// <summary>
        /// SupportByVersion OWC10 1
        /// Get
        /// </summary>
        [SupportByVersion("OWC10", 1)]
        public virtual Int32 Row
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteInt32PropertyGet(this, "Row");
            }
        }

        /// <summary>
        /// SupportByVersion OWC10 1
        /// Get/Set
        /// </summary>
        [SupportByVersion("OWC10", 1)]
        public virtual object RowHeight
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "RowHeight");
            }
            set
            {
                InvokerService.InvokeInternal.ExecuteVariantPropertySet(this, "RowHeight", value);
            }
        }

        /// <summary>
        /// SupportByVersion OWC10 1
        /// Get
        /// </summary>
        [SupportByVersion("OWC10", 1)]
        [BaseResult]
        public virtual NetOffice.OWC10Api._Range Rows
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteBaseReferencePropertyGet<NetOffice.OWC10Api._Range>(this, "Rows");
            }
        }

        /// <summary>
        /// SupportByVersion OWC10 1
        /// Get
        /// </summary>
        [SupportByVersion("OWC10", 1)]
        public virtual object Text
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "Text");
            }
        }

        /// <summary>
        /// SupportByVersion OWC10 1
        /// Get
        /// </summary>
        [SupportByVersion("OWC10", 1)]
        public virtual object Top
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "Top");
            }
        }

        /// <summary>
        /// SupportByVersion OWC10 1
        /// Get/Set
        /// </summary>
        [SupportByVersion("OWC10", 1)]
        public virtual object UseStandardHeight
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "UseStandardHeight");
            }
            set
            {
                InvokerService.InvokeInternal.ExecuteVariantPropertySet(this, "UseStandardHeight", value);
            }
        }

        /// <summary>
        /// SupportByVersion OWC10 1
        /// Get/Set
        /// </summary>
        [SupportByVersion("OWC10", 1)]
        public virtual object UseStandardWidth
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "UseStandardWidth");
            }
            set
            {
                InvokerService.InvokeInternal.ExecuteVariantPropertySet(this, "UseStandardWidth", value);
            }
        }

        /// <summary>
        /// SupportByVersion OWC10 1
        /// Get/Set
        /// </summary>
        [SupportByVersion("OWC10", 1)]
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
        /// SupportByVersion OWC10 1
        /// Get/Set
        /// </summary>
        [SupportByVersion("OWC10", 1)]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        public virtual object Value2
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "Value2");
            }
            set
            {
                InvokerService.InvokeInternal.ExecuteVariantPropertySet(this, "Value2", value);
            }
        }

        /// <summary>
        /// SupportByVersion OWC10 1
        /// Get/Set
        /// </summary>
        [SupportByVersion("OWC10", 1)]
        public virtual object VerticalAlignment
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "VerticalAlignment");
            }
            set
            {
                InvokerService.InvokeInternal.ExecuteVariantPropertySet(this, "VerticalAlignment", value);
            }
        }

        /// <summary>
        /// SupportByVersion OWC10 1
        /// Get
        /// </summary>
        [SupportByVersion("OWC10", 1)]
        public virtual object Width
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "Width");
            }
        }

        /// <summary>
        /// SupportByVersion OWC10 1
        /// Get
        /// </summary>
        [SupportByVersion("OWC10", 1)]
        public virtual NetOffice.OWC10Api.Worksheet Worksheet
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.OWC10Api.Worksheet>(this, "Worksheet", typeof(NetOffice.OWC10Api.Worksheet));
            }
        }

        #endregion

        #region Methods

        /// <summary>
        /// SupportByVersion OWC10 1
        /// </summary>
        [SupportByVersion("OWC10", 1)]
        public virtual void Activate()
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "Activate");
        }

        /// <summary>
        /// SupportByVersion OWC10 1
        /// </summary>
        /// <param name="field">optional object field</param>
        /// <param name="criteria1">optional object criteria1</param>
        /// <param name="_operator">optional object operator</param>
        /// <param name="criteria2">optional object criteria2</param>
        /// <param name="visibleDropDown">optional object visibleDropDown</param>
        [SupportByVersion("OWC10", 1)]
        public virtual void AutoFilter(object field, object criteria1, object _operator, object criteria2, object visibleDropDown)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "AutoFilter", new object[] { field, criteria1, _operator, criteria2, visibleDropDown });
        }

        /// <summary>
        /// SupportByVersion OWC10 1
        /// </summary>
        [CustomMethod]
        [SupportByVersion("OWC10", 1)]
        public virtual void AutoFilter()
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "AutoFilter");
        }

        /// <summary>
        /// SupportByVersion OWC10 1
        /// </summary>
        /// <param name="field">optional object field</param>
        [CustomMethod]
        [SupportByVersion("OWC10", 1)]
        public virtual void AutoFilter(object field)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "AutoFilter", field);
        }

        /// <summary>
        /// SupportByVersion OWC10 1
        /// </summary>
        /// <param name="field">optional object field</param>
        /// <param name="criteria1">optional object criteria1</param>
        [CustomMethod]
        [SupportByVersion("OWC10", 1)]
        public virtual void AutoFilter(object field, object criteria1)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "AutoFilter", field, criteria1);
        }

        /// <summary>
        /// SupportByVersion OWC10 1
        /// </summary>
        /// <param name="field">optional object field</param>
        /// <param name="criteria1">optional object criteria1</param>
        /// <param name="_operator">optional object operator</param>
        [CustomMethod]
        [SupportByVersion("OWC10", 1)]
        public virtual void AutoFilter(object field, object criteria1, object _operator)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "AutoFilter", field, criteria1, _operator);
        }

        /// <summary>
        /// SupportByVersion OWC10 1
        /// </summary>
        /// <param name="field">optional object field</param>
        /// <param name="criteria1">optional object criteria1</param>
        /// <param name="_operator">optional object operator</param>
        /// <param name="criteria2">optional object criteria2</param>
        [CustomMethod]
        [SupportByVersion("OWC10", 1)]
        public virtual void AutoFilter(object field, object criteria1, object _operator, object criteria2)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "AutoFilter", field, criteria1, _operator, criteria2);
        }

        /// <summary>
        /// SupportByVersion OWC10 1
        /// </summary>
        [SupportByVersion("OWC10", 1)]
        public virtual void AutoFit()
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "AutoFit");
        }

        /// <summary>
        /// SupportByVersion OWC10 1
        /// </summary>
        /// <param name="lineStyle">optional object lineStyle</param>
        /// <param name="weight">optional NetOffice.OWC10Api.Enums.XlBorderWeight Weight = 2</param>
        /// <param name="colorIndex">optional NetOffice.OWC10Api.Enums.XlColorIndex ColorIndex = -4105</param>
        /// <param name="color">optional object color</param>
        [SupportByVersion("OWC10", 1)]
        public virtual void BorderAround(object lineStyle, object weight, object colorIndex, object color)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "BorderAround", lineStyle, weight, colorIndex, color);
        }

        /// <summary>
        /// SupportByVersion OWC10 1
        /// </summary>
        [CustomMethod]
        [SupportByVersion("OWC10", 1)]
        public virtual void BorderAround()
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "BorderAround");
        }

        /// <summary>
        /// SupportByVersion OWC10 1
        /// </summary>
        /// <param name="lineStyle">optional object lineStyle</param>
        [CustomMethod]
        [SupportByVersion("OWC10", 1)]
        public virtual void BorderAround(object lineStyle)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "BorderAround", lineStyle);
        }

        /// <summary>
        /// SupportByVersion OWC10 1
        /// </summary>
        /// <param name="lineStyle">optional object lineStyle</param>
        /// <param name="weight">optional NetOffice.OWC10Api.Enums.XlBorderWeight Weight = 2</param>
        [CustomMethod]
        [SupportByVersion("OWC10", 1)]
        public virtual void BorderAround(object lineStyle, object weight)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "BorderAround", lineStyle, weight);
        }

        /// <summary>
        /// SupportByVersion OWC10 1
        /// </summary>
        /// <param name="lineStyle">optional object lineStyle</param>
        /// <param name="weight">optional NetOffice.OWC10Api.Enums.XlBorderWeight Weight = 2</param>
        /// <param name="colorIndex">optional NetOffice.OWC10Api.Enums.XlColorIndex ColorIndex = -4105</param>
        [CustomMethod]
        [SupportByVersion("OWC10", 1)]
        public virtual void BorderAround(object lineStyle, object weight, object colorIndex)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "BorderAround", lineStyle, weight, colorIndex);
        }

        /// <summary>
        /// SupportByVersion OWC10 1
        /// </summary>
        [SupportByVersion("OWC10", 1)]
        public virtual void Calculate()
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "Calculate");
        }

        /// <summary>
        /// SupportByVersion OWC10 1
        /// </summary>
        [SupportByVersion("OWC10", 1)]
        public virtual void Clear()
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "Clear");
        }

        /// <summary>
        /// SupportByVersion OWC10 1
        /// </summary>
        [SupportByVersion("OWC10", 1)]
        public virtual void ClearFormats()
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "ClearFormats");
        }

        /// <summary>
        /// SupportByVersion OWC10 1
        /// </summary>
        [SupportByVersion("OWC10", 1)]
        public virtual void ClearContents()
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "ClearContents");
        }

        /// <summary>
        /// SupportByVersion OWC10 1
        /// </summary>
        /// <param name="destination">optional object destination</param>
        [SupportByVersion("OWC10", 1)]
        public virtual void Copy(object destination)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "Copy", destination);
        }

        /// <summary>
        /// SupportByVersion OWC10 1
        /// </summary>
        [CustomMethod]
        [SupportByVersion("OWC10", 1)]
        public virtual void Copy()
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "Copy");
        }

        /// <summary>
        /// SupportByVersion OWC10 1
        /// </summary>
        /// <param name="data">object data</param>
        /// <param name="maxRows">optional object maxRows</param>
        /// <param name="maxColumns">optional object maxColumns</param>
        [SupportByVersion("OWC10", 1)]
        public virtual Int32 CopyFromRecordset(object data, object maxRows, object maxColumns)
        {
            return InvokerService.InvokeInternal.ExecuteInt32MethodGet(this, "CopyFromRecordset", data, maxRows, maxColumns);
        }

        /// <summary>
        /// SupportByVersion OWC10 1
        /// </summary>
        /// <param name="data">object data</param>
        [CustomMethod]
        [SupportByVersion("OWC10", 1)]
        public virtual Int32 CopyFromRecordset(object data)
        {
            return InvokerService.InvokeInternal.ExecuteInt32MethodGet(this, "CopyFromRecordset", data);
        }

        /// <summary>
        /// SupportByVersion OWC10 1
        /// </summary>
        /// <param name="data">object data</param>
        /// <param name="maxRows">optional object maxRows</param>
        [CustomMethod]
        [SupportByVersion("OWC10", 1)]
        public virtual Int32 CopyFromRecordset(object data, object maxRows)
        {
            return InvokerService.InvokeInternal.ExecuteInt32MethodGet(this, "CopyFromRecordset", data, maxRows);
        }

        /// <summary>
        /// SupportByVersion OWC10 1
        /// </summary>
        /// <param name="destination">optional object destination</param>
        [SupportByVersion("OWC10", 1)]
        public virtual void Cut(object destination)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "Cut", destination);
        }

        /// <summary>
        /// SupportByVersion OWC10 1
        /// </summary>
        [CustomMethod]
        [SupportByVersion("OWC10", 1)]
        public virtual void Cut()
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "Cut");
        }

        /// <summary>
        /// SupportByVersion OWC10 1
        /// </summary>
        /// <param name="shift">optional object shift</param>
        [SupportByVersion("OWC10", 1)]
        public virtual void Delete(object shift)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "Delete", shift);
        }

        /// <summary>
        /// SupportByVersion OWC10 1
        /// </summary>
        [CustomMethod]
        [SupportByVersion("OWC10", 1)]
        public virtual void Delete()
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "Delete");
        }

        /// <summary>
        /// SupportByVersion OWC10 1
        /// </summary>
        [SupportByVersion("OWC10", 1)]
        public virtual void FillDown()
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "FillDown");
        }

        /// <summary>
        /// SupportByVersion OWC10 1
        /// </summary>
        [SupportByVersion("OWC10", 1)]
        public virtual void FillRight()
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "FillRight");
        }

        /// <summary>
        /// SupportByVersion OWC10 1
        /// </summary>
        /// <param name="what">object what</param>
        /// <param name="after">optional object after</param>
        /// <param name="lookIn">optional object lookIn</param>
        /// <param name="lookAt">optional object lookAt</param>
        /// <param name="searchOrder">optional object searchOrder</param>
        /// <param name="searchDirection">optional NetOffice.OWC10Api.Enums.XlSearchDirection SearchDirection = 1</param>
        /// <param name="matchCase">optional object matchCase</param>
        /// <param name="matchByte">optional object matchByte</param>
        [SupportByVersion("OWC10", 1)]
        [BaseResult]
        public virtual NetOffice.OWC10Api._Range Find(object what, object after, object lookIn, object lookAt, object searchOrder, object searchDirection, object matchCase, object matchByte)
        {
            return InvokerService.InvokeInternal.ExecuteBaseReferenceMethodGet<NetOffice.OWC10Api._Range>(this, "Find", new object[] { what, after, lookIn, lookAt, searchOrder, searchDirection, matchCase, matchByte });
        }

        /// <summary>
        /// SupportByVersion OWC10 1
        /// </summary>
        /// <param name="what">object what</param>
        [CustomMethod]
        [BaseResult]
        [SupportByVersion("OWC10", 1)]
        public virtual NetOffice.OWC10Api._Range Find(object what)
        {
            return InvokerService.InvokeInternal.ExecuteBaseReferenceMethodGet<NetOffice.OWC10Api._Range>(this, "Find", what);
        }

        /// <summary>
        /// SupportByVersion OWC10 1
        /// </summary>
        /// <param name="what">object what</param>
        /// <param name="after">optional object after</param>
        [CustomMethod]
        [BaseResult]
        [SupportByVersion("OWC10", 1)]
        public virtual NetOffice.OWC10Api._Range Find(object what, object after)
        {
            return InvokerService.InvokeInternal.ExecuteBaseReferenceMethodGet<NetOffice.OWC10Api._Range>(this, "Find", what, after);
        }

        /// <summary>
        /// SupportByVersion OWC10 1
        /// </summary>
        /// <param name="what">object what</param>
        /// <param name="after">optional object after</param>
        /// <param name="lookIn">optional object lookIn</param>
        [CustomMethod]
        [BaseResult]
        [SupportByVersion("OWC10", 1)]
        public virtual NetOffice.OWC10Api._Range Find(object what, object after, object lookIn)
        {
            return InvokerService.InvokeInternal.ExecuteBaseReferenceMethodGet<NetOffice.OWC10Api._Range>(this, "Find", what, after, lookIn);
        }

        /// <summary>
        /// SupportByVersion OWC10 1
        /// </summary>
        /// <param name="what">object what</param>
        /// <param name="after">optional object after</param>
        /// <param name="lookIn">optional object lookIn</param>
        /// <param name="lookAt">optional object lookAt</param>
        [CustomMethod]
        [BaseResult]
        [SupportByVersion("OWC10", 1)]
        public virtual NetOffice.OWC10Api._Range Find(object what, object after, object lookIn, object lookAt)
        {
            return InvokerService.InvokeInternal.ExecuteBaseReferenceMethodGet<NetOffice.OWC10Api._Range>(this, "Find", what, after, lookIn, lookAt);
        }

        /// <summary>
        /// SupportByVersion OWC10 1
        /// </summary>
        /// <param name="what">object what</param>
        /// <param name="after">optional object after</param>
        /// <param name="lookIn">optional object lookIn</param>
        /// <param name="lookAt">optional object lookAt</param>
        /// <param name="searchOrder">optional object searchOrder</param>
        [CustomMethod]
        [BaseResult]
        [SupportByVersion("OWC10", 1)]
        public virtual NetOffice.OWC10Api._Range Find(object what, object after, object lookIn, object lookAt, object searchOrder)
        {
            return InvokerService.InvokeInternal.ExecuteBaseReferenceMethodGet<NetOffice.OWC10Api._Range>(this, "Find", new object[] { what, after, lookIn, lookAt, searchOrder });
        }

        /// <summary>
        /// SupportByVersion OWC10 1
        /// </summary>
        /// <param name="what">object what</param>
        /// <param name="after">optional object after</param>
        /// <param name="lookIn">optional object lookIn</param>
        /// <param name="lookAt">optional object lookAt</param>
        /// <param name="searchOrder">optional object searchOrder</param>
        /// <param name="searchDirection">optional NetOffice.OWC10Api.Enums.XlSearchDirection SearchDirection = 1</param>
        [CustomMethod]
        [BaseResult]
        [SupportByVersion("OWC10", 1)]
        public virtual NetOffice.OWC10Api._Range Find(object what, object after, object lookIn, object lookAt, object searchOrder, object searchDirection)
        {
            return InvokerService.InvokeInternal.ExecuteBaseReferenceMethodGet<NetOffice.OWC10Api._Range>(this, "Find", new object[] { what, after, lookIn, lookAt, searchOrder, searchDirection });
        }

        /// <summary>
        /// SupportByVersion OWC10 1
        /// </summary>
        /// <param name="what">object what</param>
        /// <param name="after">optional object after</param>
        /// <param name="lookIn">optional object lookIn</param>
        /// <param name="lookAt">optional object lookAt</param>
        /// <param name="searchOrder">optional object searchOrder</param>
        /// <param name="searchDirection">optional NetOffice.OWC10Api.Enums.XlSearchDirection SearchDirection = 1</param>
        /// <param name="matchCase">optional object matchCase</param>
        [CustomMethod]
        [BaseResult]
        [SupportByVersion("OWC10", 1)]
        public virtual NetOffice.OWC10Api._Range Find(object what, object after, object lookIn, object lookAt, object searchOrder, object searchDirection, object matchCase)
        {
            return InvokerService.InvokeInternal.ExecuteBaseReferenceMethodGet<NetOffice.OWC10Api._Range>(this, "Find", new object[] { what, after, lookIn, lookAt, searchOrder, searchDirection, matchCase });
        }

        /// <summary>
        /// SupportByVersion OWC10 1
        /// </summary>
        /// <param name="after">optional object after</param>
        [SupportByVersion("OWC10", 1)]
        [BaseResult]
        public virtual NetOffice.OWC10Api._Range FindNext(object after)
        {
            return InvokerService.InvokeInternal.ExecuteBaseReferenceMethodGet<NetOffice.OWC10Api._Range>(this, "FindNext", after);
        }

        /// <summary>
        /// SupportByVersion OWC10 1
        /// </summary>
        [CustomMethod]
        [BaseResult]
        [SupportByVersion("OWC10", 1)]
        public virtual NetOffice.OWC10Api._Range FindNext()
        {
            return InvokerService.InvokeInternal.ExecuteBaseReferenceMethodGet<NetOffice.OWC10Api._Range>(this, "FindNext");
        }

        /// <summary>
        /// SupportByVersion OWC10 1
        /// </summary>
        /// <param name="after">optional object after</param>
        [SupportByVersion("OWC10", 1)]
        [BaseResult]
        public virtual NetOffice.OWC10Api._Range FindPrevious(object after)
        {
            return InvokerService.InvokeInternal.ExecuteBaseReferenceMethodGet<NetOffice.OWC10Api._Range>(this, "FindPrevious", after);
        }

        /// <summary>
        /// SupportByVersion OWC10 1
        /// </summary>
        [CustomMethod]
        [BaseResult]
        [SupportByVersion("OWC10", 1)]
        public virtual NetOffice.OWC10Api._Range FindPrevious()
        {
            return InvokerService.InvokeInternal.ExecuteBaseReferenceMethodGet<NetOffice.OWC10Api._Range>(this, "FindPrevious");
        }

        /// <summary>
        /// SupportByVersion OWC10 1
        /// </summary>
        /// <param name="shift">optional object shift</param>
        [SupportByVersion("OWC10", 1)]
        public virtual void Insert(object shift)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "Insert", shift);
        }

        /// <summary>
        /// SupportByVersion OWC10 1
        /// </summary>
        [CustomMethod]
        [SupportByVersion("OWC10", 1)]
        public virtual void Insert()
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "Insert");
        }

        /// <summary>
        /// SupportByVersion OWC10 1
        /// </summary>
        /// <param name="file">string file</param>
        /// <param name="delimiters">optional string Delimiters = </param>
        /// <param name="consecutiveDelimAsOne">optional bool ConsecutiveDelimAsOne = false</param>
        /// <param name="textQualifier">optional string TextQualifier = \042</param>
        [SupportByVersion("OWC10", 1)]
        public virtual void LoadText(string file, object delimiters, object consecutiveDelimAsOne, object textQualifier)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "LoadText", file, delimiters, consecutiveDelimAsOne, textQualifier);
        }

        /// <summary>
        /// SupportByVersion OWC10 1
        /// </summary>
        /// <param name="file">string file</param>
        [CustomMethod]
        [SupportByVersion("OWC10", 1)]
        public virtual void LoadText(string file)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "LoadText", file);
        }

        /// <summary>
        /// SupportByVersion OWC10 1
        /// </summary>
        /// <param name="file">string file</param>
        /// <param name="delimiters">optional string Delimiters = </param>
        [CustomMethod]
        [SupportByVersion("OWC10", 1)]
        public virtual void LoadText(string file, object delimiters)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "LoadText", file, delimiters);
        }

        /// <summary>
        /// SupportByVersion OWC10 1
        /// </summary>
        /// <param name="file">string file</param>
        /// <param name="delimiters">optional string Delimiters = </param>
        /// <param name="consecutiveDelimAsOne">optional bool ConsecutiveDelimAsOne = false</param>
        [CustomMethod]
        [SupportByVersion("OWC10", 1)]
        public virtual void LoadText(string file, object delimiters, object consecutiveDelimAsOne)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "LoadText", file, delimiters, consecutiveDelimAsOne);
        }

        /// <summary>
        /// SupportByVersion OWC10 1
        /// </summary>
        /// <param name="across">optional object across</param>
        [SupportByVersion("OWC10", 1)]
        public virtual void Merge(object across)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "Merge", across);
        }

        /// <summary>
        /// SupportByVersion OWC10 1
        /// </summary>
        [CustomMethod]
        [SupportByVersion("OWC10", 1)]
        public virtual void Merge()
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "Merge");
        }

        /// <summary>
        /// SupportByVersion OWC10 1
        /// </summary>
        /// <param name="text">string text</param>
        /// <param name="delimiters">optional string Delimiters = </param>
        /// <param name="consecutiveDelimAsOne">optional bool ConsecutiveDelimAsOne = false</param>
        /// <param name="textQualifier">optional string TextQualifier = \042</param>
        [SupportByVersion("OWC10", 1)]
        public virtual void ParseText(string text, object delimiters, object consecutiveDelimAsOne, object textQualifier)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "ParseText", text, delimiters, consecutiveDelimAsOne, textQualifier);
        }

        /// <summary>
        /// SupportByVersion OWC10 1
        /// </summary>
        /// <param name="text">string text</param>
        [CustomMethod]
        [SupportByVersion("OWC10", 1)]
        public virtual void ParseText(string text)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "ParseText", text);
        }

        /// <summary>
        /// SupportByVersion OWC10 1
        /// </summary>
        /// <param name="text">string text</param>
        /// <param name="delimiters">optional string Delimiters = </param>
        [CustomMethod]
        [SupportByVersion("OWC10", 1)]
        public virtual void ParseText(string text, object delimiters)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "ParseText", text, delimiters);
        }

        /// <summary>
        /// SupportByVersion OWC10 1
        /// </summary>
        /// <param name="text">string text</param>
        /// <param name="delimiters">optional string Delimiters = </param>
        /// <param name="consecutiveDelimAsOne">optional bool ConsecutiveDelimAsOne = false</param>
        [CustomMethod]
        [SupportByVersion("OWC10", 1)]
        public virtual void ParseText(string text, object delimiters, object consecutiveDelimAsOne)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "ParseText", text, delimiters, consecutiveDelimAsOne);
        }

        /// <summary>
        /// SupportByVersion OWC10 1
        /// </summary>
        [SupportByVersion("OWC10", 1)]
        public virtual void Paste()
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "Paste");
        }

        /// <summary>
        /// SupportByVersion OWC10 1
        /// </summary>
        [SupportByVersion("OWC10", 1)]
        public virtual void Select()
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "Select");
        }

        /// <summary>
        /// SupportByVersion OWC10 1
        /// </summary>
        [SupportByVersion("OWC10", 1)]
        public virtual void Show()
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "Show");
        }

        /// <summary>
        /// SupportByVersion OWC10 1
        /// </summary>
        /// <param name="columnKey">optional Int32 ColumnKey = 1</param>
        /// <param name="order">optional NetOffice.OWC10Api.Enums.XlSortOrder Order = 1</param>
        /// <param name="header">optional NetOffice.OWC10Api.Enums.XlYesNoGuess Header = 2</param>
        [SupportByVersion("OWC10", 1)]
        public virtual void Sort(object columnKey, object order, object header)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "Sort", columnKey, order, header);
        }

        /// <summary>
        /// SupportByVersion OWC10 1
        /// </summary>
        [CustomMethod]
        [SupportByVersion("OWC10", 1)]
        public virtual void Sort()
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "Sort");
        }

        /// <summary>
        /// SupportByVersion OWC10 1
        /// </summary>
        /// <param name="columnKey">optional Int32 ColumnKey = 1</param>
        [CustomMethod]
        [SupportByVersion("OWC10", 1)]
        public virtual void Sort(object columnKey)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "Sort", columnKey);
        }

        /// <summary>
        /// SupportByVersion OWC10 1
        /// </summary>
        /// <param name="columnKey">optional Int32 ColumnKey = 1</param>
        /// <param name="order">optional NetOffice.OWC10Api.Enums.XlSortOrder Order = 1</param>
        [CustomMethod]
        [SupportByVersion("OWC10", 1)]
        public virtual void Sort(object columnKey, object order)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "Sort", columnKey, order);
        }

        /// <summary>
        /// SupportByVersion OWC10 1
        /// </summary>
        [SupportByVersion("OWC10", 1)]
        public virtual void UnMerge()
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "UnMerge");
        }

        #endregion

        #region IEnumerableProvider<object>

        ICOMObject IEnumerableProvider<object>.GetComObjectEnumerator(ICOMObject parent)
        {
            return NetOffice.Utils.GetComObjectEnumeratorAsProperty(parent, this, false);
        }

        IEnumerable IEnumerableProvider<object>.FetchVariantComObjectEnumerator(ICOMObject parent, ICOMObject enumerator)
        {
            return NetOffice.Utils.FetchVariantComObjectEnumerator(parent, enumerator, true);
        }

        #endregion

        #region IEnumerable<object>

        /// <summary>
        /// SupportByVersion OWC10, 1
        /// </summary>
        [SupportByVersion("OWC10", 1)]
        public virtual IEnumerator<object> GetEnumerator()
        {
            NetRuntimeSystem.Collections.IEnumerable innerEnumerator = (this as NetRuntimeSystem.Collections.IEnumerable);
            foreach (object item in innerEnumerator)
                yield return item;
        }

        #endregion

        #region IEnumerable

        /// <summary>
        /// SupportByVersion OWC10, 1
        /// </summary>
        [SupportByVersion("OWC10", 1)]
        IEnumerator NetRuntimeSystem.Collections.IEnumerable.GetEnumerator()
        {
            return NetOffice.Utils.GetProxyEnumeratorAsProperty(this, true);
        }

        #endregion

        #pragma warning restore
    }
}

