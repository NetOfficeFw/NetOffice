using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;
using NetOffice.ExcelApi;

namespace NetOffice.ExcelApi.Behind
{
    /// <summary>
    /// IPivotTable
    /// </summary>
    [SyntaxBypass]
    public class IPivotTable_ : COMObject, NetOffice.ExcelApi.IPivotTable_
    {
        #region Ctor

        /// <summary>
        /// Stub Ctor, not indented to use
        /// </summary>
        public IPivotTable_() : base()
        {
        }

        #endregion

        #region Properties

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// Unknown COM Proxy
        /// </summary>
        /// <param name="index">optional object index</param>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16), ProxyResult]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        public virtual object get_ColumnFields(object index)
        {
            return InvokerService.InvokeInternal.ExecuteReferencePropertyGet(this, "ColumnFields", index);
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Alias for get_ColumnFields
        /// Unknown COM Proxy
        /// </summary>
        /// <param name="index">optional object index</param>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16), ProxyResult, Redirect("get_ColumnFields")]
        public virtual object ColumnFields(object index)
        {
            return get_ColumnFields(index);
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// Unknown COM Proxy
        /// </summary>
        /// <param name="index">optional object index</param>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16), ProxyResult]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        public virtual object get_DataFields(object index)
        {
            return InvokerService.InvokeInternal.ExecuteReferencePropertyGet(this, "DataFields", index);
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Alias for get_DataFields
        /// Unknown COM Proxy
        /// </summary>
        /// <param name="index">optional object index</param>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16), ProxyResult, Redirect("get_DataFields")]
        public virtual object DataFields(object index)
        {
            return get_DataFields(index);
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// Unknown COM Proxy
        /// </summary>
        /// <param name="index">optional object index</param>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16), ProxyResult]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        public virtual object get_HiddenFields(object index)
        {
            return InvokerService.InvokeInternal.ExecuteReferencePropertyGet(this, "HiddenFields", index);
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Alias for get_HiddenFields
        /// Unknown COM Proxy
        /// </summary>
        /// <param name="index">optional object index</param>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16), ProxyResult, Redirect("get_HiddenFields")]
        public virtual object HiddenFields(object index)
        {
            return get_HiddenFields(index);
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// Unknown COM Proxy
        /// </summary>
        /// <param name="index">optional object index</param>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16), ProxyResult]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        public virtual object get_PageFields(object index)
        {
            return InvokerService.InvokeInternal.ExecuteReferencePropertyGet(this, "PageFields", index);
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Alias for get_PageFields
        /// Unknown COM Proxy
        /// </summary>
        /// <param name="index">optional object index</param>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16), ProxyResult, Redirect("get_PageFields")]
        public virtual object PageFields(object index)
        {
            return get_PageFields(index);
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// Unknown COM Proxy
        /// </summary>
        /// <param name="index">optional object index</param>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16), ProxyResult]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        public virtual object get_RowFields(object index)
        {
            return InvokerService.InvokeInternal.ExecuteReferencePropertyGet(this, "RowFields", index);
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Alias for get_RowFields
        /// Unknown COM Proxy
        /// </summary>
        /// <param name="index">optional object index</param>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16), ProxyResult, Redirect("get_RowFields")]
        public virtual object RowFields(object index)
        {
            return get_RowFields(index);
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// Unknown COM Proxy
        /// </summary>
        /// <param name="index">optional object index</param>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16), ProxyResult]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        public virtual object get_VisibleFields(object index)
        {
            return InvokerService.InvokeInternal.ExecuteReferencePropertyGet(this, "VisibleFields", index);
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Alias for get_VisibleFields
        /// Unknown COM Proxy
        /// </summary>
        /// <param name="index">optional object index</param>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16), ProxyResult, Redirect("get_VisibleFields")]
        public virtual object VisibleFields(object index)
        {
            return get_VisibleFields(index);
        }

        #endregion

        #region Methods

        #endregion
    }

    /// <summary>
    /// Interface IPivotTable 
    /// SupportByVersion Excel, 9,10,11,12,14,15,16
    /// </summary>
    [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
    [EntityType(EntityType.IsInterface)]
    public class IPivotTable : NetOffice.ExcelApi.Behind.IPivotTable_, NetOffice.ExcelApi.IPivotTable_
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
                    _contractType = typeof(NetOffice.ExcelApi.IPivotTable);
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
                    _type = typeof(IPivotTable);
                return _type;
            }
        }

        #endregion

        #region Ctor

        /// <summary>
        /// Stub Ctor, not indented to use
        /// </summary>
        public IPivotTable() : base()
        {
        }

        #endregion

        #region Properties

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual NetOffice.ExcelApi.Application Application
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.ExcelApi.Application>(this, "Application", typeof(NetOffice.ExcelApi.Application));
            }
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual NetOffice.ExcelApi.Enums.XlCreator Creator
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteEnumPropertyGet<NetOffice.ExcelApi.Enums.XlCreator>(this, "Creator");
            }
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// Unknown COM Proxy
        /// </summary>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16), ProxyResult]
        public virtual object Parent
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteReferencePropertyGet(this, "Parent");
            }
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// Unknown COM Proxy
        /// </summary>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16), ProxyResult]
        public virtual object ColumnFields
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteReferencePropertyGet(this, "ColumnFields");
            }
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual bool ColumnGrand
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "ColumnGrand");
            }
            set
            {
                InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "ColumnGrand", value);
            }
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual NetOffice.ExcelApi.Range ColumnRange
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.ExcelApi.Range>(this, "ColumnRange", typeof(NetOffice.ExcelApi.Range));
            }
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual NetOffice.ExcelApi.Range DataBodyRange
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.ExcelApi.Range>(this, "DataBodyRange", typeof(NetOffice.ExcelApi.Range));
            }
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// Unknown COM Proxy
        /// </summary>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16), ProxyResult]
        public virtual object DataFields
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteReferencePropertyGet(this, "DataFields");
            }
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual NetOffice.ExcelApi.Range DataLabelRange
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.ExcelApi.Range>(this, "DataLabelRange", typeof(NetOffice.ExcelApi.Range));
            }
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual string _Default
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "_Default");
            }
            set
            {
                InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "_Default", value);
            }
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual bool HasAutoFormat
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "HasAutoFormat");
            }
            set
            {
                InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "HasAutoFormat", value);
            }
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// Unknown COM Proxy
        /// </summary>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16), ProxyResult]
        public virtual object HiddenFields
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteReferencePropertyGet(this, "HiddenFields");
            }
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual string InnerDetail
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "InnerDetail");
            }
            set
            {
                InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "InnerDetail", value);
            }
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
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
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// Unknown COM Proxy
        /// </summary>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16), ProxyResult]
        public virtual object PageFields
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteReferencePropertyGet(this, "PageFields");
            }
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual NetOffice.ExcelApi.Range PageRange
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.ExcelApi.Range>(this, "PageRange", typeof(NetOffice.ExcelApi.Range));
            }
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual NetOffice.ExcelApi.Range PageRangeCells
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.ExcelApi.Range>(this, "PageRangeCells", typeof(NetOffice.ExcelApi.Range));
            }
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual DateTime RefreshDate
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteDateTimePropertyGet(this, "RefreshDate");
            }
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual string RefreshName
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "RefreshName");
            }
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// Unknown COM Proxy
        /// </summary>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16), ProxyResult]
        public virtual object RowFields
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteReferencePropertyGet(this, "RowFields");
            }
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual bool RowGrand
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "RowGrand");
            }
            set
            {
                InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "RowGrand", value);
            }
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual NetOffice.ExcelApi.Range RowRange
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.ExcelApi.Range>(this, "RowRange", typeof(NetOffice.ExcelApi.Range));
            }
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual bool SaveData
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "SaveData");
            }
            set
            {
                InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "SaveData", value);
            }
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual object SourceData
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "SourceData");
            }
            set
            {
                InvokerService.InvokeInternal.ExecuteVariantPropertySet(this, "SourceData", value);
            }
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual NetOffice.ExcelApi.Range TableRange1
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.ExcelApi.Range>(this, "TableRange1", typeof(NetOffice.ExcelApi.Range));
            }
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual NetOffice.ExcelApi.Range TableRange2
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.ExcelApi.Range>(this, "TableRange2", typeof(NetOffice.ExcelApi.Range));
            }
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual string Value
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "Value");
            }
            set
            {
                InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "Value", value);
            }
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// Unknown COM Proxy
        /// </summary>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16), ProxyResult]
        public virtual object VisibleFields
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteReferencePropertyGet(this, "VisibleFields");
            }
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual Int32 CacheIndex
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteInt32PropertyGet(this, "CacheIndex");
            }
            set
            {
                InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "CacheIndex", value);
            }
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual bool DisplayErrorString
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "DisplayErrorString");
            }
            set
            {
                InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "DisplayErrorString", value);
            }
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual bool DisplayNullString
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "DisplayNullString");
            }
            set
            {
                InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "DisplayNullString", value);
            }
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual bool EnableDrilldown
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "EnableDrilldown");
            }
            set
            {
                InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "EnableDrilldown", value);
            }
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual bool EnableFieldDialog
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "EnableFieldDialog");
            }
            set
            {
                InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "EnableFieldDialog", value);
            }
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual bool EnableWizard
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "EnableWizard");
            }
            set
            {
                InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "EnableWizard", value);
            }
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual string ErrorString
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "ErrorString");
            }
            set
            {
                InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "ErrorString", value);
            }
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual bool ManualUpdate
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "ManualUpdate");
            }
            set
            {
                InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "ManualUpdate", value);
            }
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual bool MergeLabels
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "MergeLabels");
            }
            set
            {
                InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "MergeLabels", value);
            }
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual string NullString
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "NullString");
            }
            set
            {
                InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "NullString", value);
            }
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual NetOffice.ExcelApi.PivotFormulas PivotFormulas
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.ExcelApi.PivotFormulas>(this, "PivotFormulas", typeof(NetOffice.ExcelApi.PivotFormulas));
            }
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual bool SubtotalHiddenPageItems
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "SubtotalHiddenPageItems");
            }
            set
            {
                InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "SubtotalHiddenPageItems", value);
            }
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual Int32 PageFieldOrder
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteInt32PropertyGet(this, "PageFieldOrder");
            }
            set
            {
                InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "PageFieldOrder", value);
            }
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual string PageFieldStyle
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "PageFieldStyle");
            }
            set
            {
                InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "PageFieldStyle", value);
            }
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual Int32 PageFieldWrapCount
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteInt32PropertyGet(this, "PageFieldWrapCount");
            }
            set
            {
                InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "PageFieldWrapCount", value);
            }
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual bool PreserveFormatting
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "PreserveFormatting");
            }
            set
            {
                InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "PreserveFormatting", value);
            }
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual string PivotSelection
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "PivotSelection");
            }
            set
            {
                InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "PivotSelection", value);
            }
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual NetOffice.ExcelApi.Enums.XlPTSelectionMode SelectionMode
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteEnumPropertyGet<NetOffice.ExcelApi.Enums.XlPTSelectionMode>(this, "SelectionMode");
            }
            set
            {
                InvokerService.InvokeInternal.ExecuteEnumPropertySet(this, "SelectionMode", value);
            }
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual string TableStyle
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "TableStyle");
            }
            set
            {
                InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "TableStyle", value);
            }
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
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
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual string VacatedStyle
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "VacatedStyle");
            }
            set
            {
                InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "VacatedStyle", value);
            }
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual bool PrintTitles
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "PrintTitles");
            }
            set
            {
                InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "PrintTitles", value);
            }
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual NetOffice.ExcelApi.CubeFields CubeFields
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.ExcelApi.CubeFields>(this, "CubeFields", typeof(NetOffice.ExcelApi.CubeFields));
            }
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual string GrandTotalName
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "GrandTotalName");
            }
            set
            {
                InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "GrandTotalName", value);
            }
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual bool SmallGrid
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "SmallGrid");
            }
            set
            {
                InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "SmallGrid", value);
            }
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual bool RepeatItemsOnEachPrintedPage
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "RepeatItemsOnEachPrintedPage");
            }
            set
            {
                InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "RepeatItemsOnEachPrintedPage", value);
            }
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual bool TotalsAnnotation
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "TotalsAnnotation");
            }
            set
            {
                InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "TotalsAnnotation", value);
            }
        }

        /// <summary>
        /// SupportByVersion Excel 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Excel", 10, 11, 12, 14, 15, 16)]
        public virtual string PivotSelectionStandard
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "PivotSelectionStandard");
            }
            set
            {
                InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "PivotSelectionStandard", value);
            }
        }

        /// <summary>
        /// SupportByVersion Excel 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Excel", 10, 11, 12, 14, 15, 16)]
        public virtual NetOffice.ExcelApi.PivotField DataPivotField
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.ExcelApi.PivotField>(this, "DataPivotField", typeof(NetOffice.ExcelApi.PivotField));
            }
        }

        /// <summary>
        /// SupportByVersion Excel 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Excel", 10, 11, 12, 14, 15, 16)]
        public virtual bool EnableDataValueEditing
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "EnableDataValueEditing");
            }
            set
            {
                InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "EnableDataValueEditing", value);
            }
        }

        /// <summary>
        /// SupportByVersion Excel 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Excel", 10, 11, 12, 14, 15, 16)]
        public virtual string MDX
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "MDX");
            }
        }

        /// <summary>
        /// SupportByVersion Excel 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Excel", 10, 11, 12, 14, 15, 16)]
        public virtual bool ViewCalculatedMembers
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "ViewCalculatedMembers");
            }
            set
            {
                InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "ViewCalculatedMembers", value);
            }
        }

        /// <summary>
        /// SupportByVersion Excel 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Excel", 10, 11, 12, 14, 15, 16)]
        public virtual NetOffice.ExcelApi.CalculatedMembers CalculatedMembers
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.ExcelApi.CalculatedMembers>(this, "CalculatedMembers", typeof(NetOffice.ExcelApi.CalculatedMembers));
            }
        }

        /// <summary>
        /// SupportByVersion Excel 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Excel", 10, 11, 12, 14, 15, 16)]
        public virtual bool DisplayImmediateItems
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "DisplayImmediateItems");
            }
            set
            {
                InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "DisplayImmediateItems", value);
            }
        }

        /// <summary>
        /// SupportByVersion Excel 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Excel", 10, 11, 12, 14, 15, 16)]
        public virtual bool EnableFieldList
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "EnableFieldList");
            }
            set
            {
                InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "EnableFieldList", value);
            }
        }

        /// <summary>
        /// SupportByVersion Excel 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Excel", 10, 11, 12, 14, 15, 16)]
        public virtual bool VisualTotals
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "VisualTotals");
            }
            set
            {
                InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "VisualTotals", value);
            }
        }

        /// <summary>
        /// SupportByVersion Excel 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Excel", 10, 11, 12, 14, 15, 16)]
        public virtual bool ShowPageMultipleItemLabel
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "ShowPageMultipleItemLabel");
            }
            set
            {
                InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "ShowPageMultipleItemLabel", value);
            }
        }

        /// <summary>
        /// SupportByVersion Excel 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Excel", 10, 11, 12, 14, 15, 16)]
        public virtual NetOffice.ExcelApi.Enums.XlPivotTableVersionList Version
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteEnumPropertyGet<NetOffice.ExcelApi.Enums.XlPivotTableVersionList>(this, "Version");
            }
        }

        /// <summary>
        /// SupportByVersion Excel 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Excel", 10, 11, 12, 14, 15, 16)]
        public virtual bool DisplayEmptyRow
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "DisplayEmptyRow");
            }
            set
            {
                InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "DisplayEmptyRow", value);
            }
        }

        /// <summary>
        /// SupportByVersion Excel 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Excel", 10, 11, 12, 14, 15, 16)]
        public virtual bool DisplayEmptyColumn
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "DisplayEmptyColumn");
            }
            set
            {
                InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "DisplayEmptyColumn", value);
            }
        }

        /// <summary>
        /// SupportByVersion Excel 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Excel", 10, 11, 12, 14, 15, 16)]
        public virtual bool ShowCellBackgroundFromOLAP
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "ShowCellBackgroundFromOLAP");
            }
            set
            {
                InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "ShowCellBackgroundFromOLAP", value);
            }
        }

        /// <summary>
        /// SupportByVersion Excel 12, 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Excel", 12, 14, 15, 16)]
        public virtual NetOffice.ExcelApi.PivotAxis PivotColumnAxis
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.ExcelApi.PivotAxis>(this, "PivotColumnAxis", typeof(NetOffice.ExcelApi.PivotAxis));
            }
        }

        /// <summary>
        /// SupportByVersion Excel 12, 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Excel", 12, 14, 15, 16)]
        public virtual NetOffice.ExcelApi.PivotAxis PivotRowAxis
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.ExcelApi.PivotAxis>(this, "PivotRowAxis", typeof(NetOffice.ExcelApi.PivotAxis));
            }
        }

        /// <summary>
        /// SupportByVersion Excel 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Excel", 12, 14, 15, 16)]
        public virtual bool ShowDrillIndicators
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "ShowDrillIndicators");
            }
            set
            {
                InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "ShowDrillIndicators", value);
            }
        }

        /// <summary>
        /// SupportByVersion Excel 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Excel", 12, 14, 15, 16)]
        public virtual bool PrintDrillIndicators
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "PrintDrillIndicators");
            }
            set
            {
                InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "PrintDrillIndicators", value);
            }
        }

        /// <summary>
        /// SupportByVersion Excel 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Excel", 12, 14, 15, 16)]
        public virtual bool DisplayMemberPropertyTooltips
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "DisplayMemberPropertyTooltips");
            }
            set
            {
                InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "DisplayMemberPropertyTooltips", value);
            }
        }

        /// <summary>
        /// SupportByVersion Excel 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Excel", 12, 14, 15, 16)]
        public virtual bool DisplayContextTooltips
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "DisplayContextTooltips");
            }
            set
            {
                InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "DisplayContextTooltips", value);
            }
        }

        /// <summary>
        /// SupportByVersion Excel 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Excel", 12, 14, 15, 16)]
        public virtual Int32 CompactRowIndent
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteInt32PropertyGet(this, "CompactRowIndent");
            }
            set
            {
                InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "CompactRowIndent", value);
            }
        }

        /// <summary>
        /// SupportByVersion Excel 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Excel", 12, 14, 15, 16)]
        public virtual NetOffice.ExcelApi.Enums.XlLayoutRowType LayoutRowDefault
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteEnumPropertyGet<NetOffice.ExcelApi.Enums.XlLayoutRowType>(this, "LayoutRowDefault");
            }
            set
            {
                InvokerService.InvokeInternal.ExecuteEnumPropertySet(this, "LayoutRowDefault", value);
            }
        }

        /// <summary>
        /// SupportByVersion Excel 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Excel", 12, 14, 15, 16)]
        public virtual bool DisplayFieldCaptions
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "DisplayFieldCaptions");
            }
            set
            {
                InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "DisplayFieldCaptions", value);
            }
        }

        /// <summary>
        /// SupportByVersion Excel 12, 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Excel", 12, 14, 15, 16)]
        public virtual NetOffice.ExcelApi.PivotFilters ActiveFilters
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.ExcelApi.PivotFilters>(this, "ActiveFilters", typeof(NetOffice.ExcelApi.PivotFilters));
            }
        }

        /// <summary>
        /// SupportByVersion Excel 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Excel", 12, 14, 15, 16)]
        public virtual bool InGridDropZones
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "InGridDropZones");
            }
            set
            {
                InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "InGridDropZones", value);
            }
        }

        /// <summary>
        /// SupportByVersion Excel 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Excel", 12, 14, 15, 16)]
        public virtual object TableStyle2
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "TableStyle2");
            }
            set
            {
                InvokerService.InvokeInternal.ExecuteVariantPropertySet(this, "TableStyle2", value);
            }
        }

        /// <summary>
        /// SupportByVersion Excel 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Excel", 12, 14, 15, 16)]
        public virtual bool ShowTableStyleLastColumn
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "ShowTableStyleLastColumn");
            }
            set
            {
                InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "ShowTableStyleLastColumn", value);
            }
        }

        /// <summary>
        /// SupportByVersion Excel 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Excel", 12, 14, 15, 16)]
        public virtual bool ShowTableStyleRowStripes
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "ShowTableStyleRowStripes");
            }
            set
            {
                InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "ShowTableStyleRowStripes", value);
            }
        }

        /// <summary>
        /// SupportByVersion Excel 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Excel", 12, 14, 15, 16)]
        public virtual bool ShowTableStyleColumnStripes
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "ShowTableStyleColumnStripes");
            }
            set
            {
                InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "ShowTableStyleColumnStripes", value);
            }
        }

        /// <summary>
        /// SupportByVersion Excel 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Excel", 12, 14, 15, 16)]
        public virtual bool ShowTableStyleRowHeaders
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "ShowTableStyleRowHeaders");
            }
            set
            {
                InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "ShowTableStyleRowHeaders", value);
            }
        }

        /// <summary>
        /// SupportByVersion Excel 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Excel", 12, 14, 15, 16)]
        public virtual bool ShowTableStyleColumnHeaders
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "ShowTableStyleColumnHeaders");
            }
            set
            {
                InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "ShowTableStyleColumnHeaders", value);
            }
        }

        /// <summary>
        /// SupportByVersion Excel 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Excel", 12, 14, 15, 16)]
        public virtual bool AllowMultipleFilters
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "AllowMultipleFilters");
            }
            set
            {
                InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "AllowMultipleFilters", value);
            }
        }

        /// <summary>
        /// SupportByVersion Excel 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Excel", 12, 14, 15, 16)]
        public virtual string CompactLayoutRowHeader
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "CompactLayoutRowHeader");
            }
            set
            {
                InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "CompactLayoutRowHeader", value);
            }
        }

        /// <summary>
        /// SupportByVersion Excel 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Excel", 12, 14, 15, 16)]
        public virtual string CompactLayoutColumnHeader
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "CompactLayoutColumnHeader");
            }
            set
            {
                InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "CompactLayoutColumnHeader", value);
            }
        }

        /// <summary>
        /// SupportByVersion Excel 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Excel", 12, 14, 15, 16)]
        public virtual bool FieldListSortAscending
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "FieldListSortAscending");
            }
            set
            {
                InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "FieldListSortAscending", value);
            }
        }

        /// <summary>
        /// SupportByVersion Excel 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Excel", 12, 14, 15, 16)]
        public virtual bool SortUsingCustomLists
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "SortUsingCustomLists");
            }
            set
            {
                InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "SortUsingCustomLists", value);
            }
        }

        /// <summary>
        /// SupportByVersion Excel 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Excel", 12, 14, 15, 16)]
        public virtual string Location
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "Location");
            }
            set
            {
                InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "Location", value);
            }
        }

        /// <summary>
        /// SupportByVersion Excel 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Excel", 14, 15, 16)]
        public virtual bool EnableWriteback
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "EnableWriteback");
            }
            set
            {
                InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "EnableWriteback", value);
            }
        }

        /// <summary>
        /// SupportByVersion Excel 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Excel", 14, 15, 16)]
        public virtual NetOffice.ExcelApi.Enums.XlAllocation Allocation
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteEnumPropertyGet<NetOffice.ExcelApi.Enums.XlAllocation>(this, "Allocation");
            }
            set
            {
                InvokerService.InvokeInternal.ExecuteEnumPropertySet(this, "Allocation", value);
            }
        }

        /// <summary>
        /// SupportByVersion Excel 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Excel", 14, 15, 16)]
        public virtual NetOffice.ExcelApi.Enums.XlAllocationValue AllocationValue
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteEnumPropertyGet<NetOffice.ExcelApi.Enums.XlAllocationValue>(this, "AllocationValue");
            }
            set
            {
                InvokerService.InvokeInternal.ExecuteEnumPropertySet(this, "AllocationValue", value);
            }
        }

        /// <summary>
        /// SupportByVersion Excel 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Excel", 14, 15, 16)]
        public virtual NetOffice.ExcelApi.Enums.XlAllocationMethod AllocationMethod
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteEnumPropertyGet<NetOffice.ExcelApi.Enums.XlAllocationMethod>(this, "AllocationMethod");
            }
            set
            {
                InvokerService.InvokeInternal.ExecuteEnumPropertySet(this, "AllocationMethod", value);
            }
        }

        /// <summary>
        /// SupportByVersion Excel 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Excel", 14, 15, 16)]
        public virtual string AllocationWeightExpression
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "AllocationWeightExpression");
            }
            set
            {
                InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "AllocationWeightExpression", value);
            }
        }

        /// <summary>
        /// SupportByVersion Excel 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Excel", 14, 15, 16)]
        public virtual NetOffice.ExcelApi.PivotTableChangeList ChangeList
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.ExcelApi.PivotTableChangeList>(this, "ChangeList", typeof(NetOffice.ExcelApi.PivotTableChangeList));
            }
        }

        /// <summary>
        /// SupportByVersion Excel 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Excel", 14, 15, 16)]
        public virtual NetOffice.ExcelApi.Slicers Slicers
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.ExcelApi.Slicers>(this, "Slicers", typeof(NetOffice.ExcelApi.Slicers));
            }
        }

        /// <summary>
        /// SupportByVersion Excel 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Excel", 14, 15, 16)]
        public virtual string AlternativeText
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "AlternativeText");
            }
            set
            {
                InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "AlternativeText", value);
            }
        }

        /// <summary>
        /// SupportByVersion Excel 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Excel", 14, 15, 16)]
        public virtual string Summary
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "Summary");
            }
            set
            {
                InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "Summary", value);
            }
        }

        /// <summary>
        /// SupportByVersion Excel 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Excel", 14, 15, 16)]
        public virtual bool VisualTotalsForSets
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "VisualTotalsForSets");
            }
            set
            {
                InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "VisualTotalsForSets", value);
            }
        }

        /// <summary>
        /// SupportByVersion Excel 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Excel", 14, 15, 16)]
        public virtual bool ShowValuesRow
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "ShowValuesRow");
            }
            set
            {
                InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "ShowValuesRow", value);
            }
        }

        /// <summary>
        /// SupportByVersion Excel 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Excel", 14, 15, 16)]
        public virtual bool CalculatedMembersInFilters
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "CalculatedMembersInFilters");
            }
            set
            {
                InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "CalculatedMembersInFilters", value);
            }
        }

        /// <summary>
        /// SupportByVersion Excel 15,16
        /// Get
        /// </summary>
        [SupportByVersion("Excel", 15, 16)]
        public virtual bool Hidden
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "Hidden");
            }
        }

        /// <summary>
        /// SupportByVersion Excel 15,16
        /// Get
        /// </summary>
        [SupportByVersion("Excel", 15, 16)]
        public virtual NetOffice.ExcelApi.Shape PivotChart
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.ExcelApi.Shape>(this, "PivotChart", typeof(NetOffice.ExcelApi.Shape));
            }
        }

        #endregion

        #region Methods

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="rowFields">optional object rowFields</param>
        /// <param name="columnFields">optional object columnFields</param>
        /// <param name="pageFields">optional object pageFields</param>
        /// <param name="addToTable">optional object addToTable</param>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual object AddFields(object rowFields, object columnFields, object pageFields, object addToTable)
        {
            return InvokerService.InvokeInternal.ExecuteVariantMethodGet(this, "AddFields", rowFields, columnFields, pageFields, addToTable);
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        [CustomMethod]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual object AddFields()
        {
            return InvokerService.InvokeInternal.ExecuteVariantMethodGet(this, "AddFields");
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="rowFields">optional object rowFields</param>
        [CustomMethod]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual object AddFields(object rowFields)
        {
            return InvokerService.InvokeInternal.ExecuteVariantMethodGet(this, "AddFields", rowFields);
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="rowFields">optional object rowFields</param>
        /// <param name="columnFields">optional object columnFields</param>
        [CustomMethod]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual object AddFields(object rowFields, object columnFields)
        {
            return InvokerService.InvokeInternal.ExecuteVariantMethodGet(this, "AddFields", rowFields, columnFields);
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="rowFields">optional object rowFields</param>
        /// <param name="columnFields">optional object columnFields</param>
        /// <param name="pageFields">optional object pageFields</param>
        [CustomMethod]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual object AddFields(object rowFields, object columnFields, object pageFields)
        {
            return InvokerService.InvokeInternal.ExecuteVariantMethodGet(this, "AddFields", rowFields, columnFields, pageFields);
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="pageField">optional object pageField</param>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual object ShowPages(object pageField)
        {
            return InvokerService.InvokeInternal.ExecuteVariantMethodGet(this, "ShowPages", pageField);
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        [CustomMethod]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual object ShowPages()
        {
            return InvokerService.InvokeInternal.ExecuteVariantMethodGet(this, "ShowPages");
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="index">optional object index</param>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual object PivotFields(object index)
        {
            return InvokerService.InvokeInternal.ExecuteVariantMethodGet(this, "PivotFields", index);
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        [CustomMethod]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual object PivotFields()
        {
            return InvokerService.InvokeInternal.ExecuteVariantMethodGet(this, "PivotFields");
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual bool RefreshTable()
        {
            return InvokerService.InvokeInternal.ExecuteBoolMethodGet(this, "RefreshTable");
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual NetOffice.ExcelApi.CalculatedFields CalculatedFields()
        {
            return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.ExcelApi.CalculatedFields>(this, "CalculatedFields", typeof(NetOffice.ExcelApi.CalculatedFields));
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="name">string name</param>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual Double GetData(string name)
        {
            return InvokerService.InvokeInternal.ExecuteDoubleMethodGet(this, "GetData", name);
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual Int32 ListFormulas()
        {
            return InvokerService.InvokeInternal.ExecuteInt32MethodGet(this, "ListFormulas");
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual NetOffice.ExcelApi.PivotCache PivotCache()
        {
            return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.ExcelApi.PivotCache>(this, "PivotCache", typeof(NetOffice.ExcelApi.PivotCache));
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="sourceType">optional object sourceType</param>
        /// <param name="sourceData">optional object sourceData</param>
        /// <param name="tableDestination">optional object tableDestination</param>
        /// <param name="tableName">optional object tableName</param>
        /// <param name="rowGrand">optional object rowGrand</param>
        /// <param name="columnGrand">optional object columnGrand</param>
        /// <param name="saveData">optional object saveData</param>
        /// <param name="hasAutoFormat">optional object hasAutoFormat</param>
        /// <param name="autoPage">optional object autoPage</param>
        /// <param name="reserved">optional object reserved</param>
        /// <param name="backgroundQuery">optional object backgroundQuery</param>
        /// <param name="optimizeCache">optional object optimizeCache</param>
        /// <param name="pageFieldOrder">optional object pageFieldOrder</param>
        /// <param name="pageFieldWrapCount">optional object pageFieldWrapCount</param>
        /// <param name="readData">optional object readData</param>
        /// <param name="connection">optional object connection</param>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual Int32 PivotTableWizard(object sourceType, object sourceData, object tableDestination, object tableName, object rowGrand, object columnGrand, object saveData, object hasAutoFormat, object autoPage, object reserved, object backgroundQuery, object optimizeCache, object pageFieldOrder, object pageFieldWrapCount, object readData, object connection)
        {
            return InvokerService.InvokeInternal.ExecuteInt32MethodGet(this, "PivotTableWizard", new object[] { sourceType, sourceData, tableDestination, tableName, rowGrand, columnGrand, saveData, hasAutoFormat, autoPage, reserved, backgroundQuery, optimizeCache, pageFieldOrder, pageFieldWrapCount, readData, connection });
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        [CustomMethod]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual Int32 PivotTableWizard()
        {
            return InvokerService.InvokeInternal.ExecuteInt32MethodGet(this, "PivotTableWizard");
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="sourceType">optional object sourceType</param>
        [CustomMethod]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual Int32 PivotTableWizard(object sourceType)
        {
            return InvokerService.InvokeInternal.ExecuteInt32MethodGet(this, "PivotTableWizard", sourceType);
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="sourceType">optional object sourceType</param>
        /// <param name="sourceData">optional object sourceData</param>
        [CustomMethod]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual Int32 PivotTableWizard(object sourceType, object sourceData)
        {
            return InvokerService.InvokeInternal.ExecuteInt32MethodGet(this, "PivotTableWizard", sourceType, sourceData);
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="sourceType">optional object sourceType</param>
        /// <param name="sourceData">optional object sourceData</param>
        /// <param name="tableDestination">optional object tableDestination</param>
        [CustomMethod]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual Int32 PivotTableWizard(object sourceType, object sourceData, object tableDestination)
        {
            return InvokerService.InvokeInternal.ExecuteInt32MethodGet(this, "PivotTableWizard", sourceType, sourceData, tableDestination);
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="sourceType">optional object sourceType</param>
        /// <param name="sourceData">optional object sourceData</param>
        /// <param name="tableDestination">optional object tableDestination</param>
        /// <param name="tableName">optional object tableName</param>
        [CustomMethod]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual Int32 PivotTableWizard(object sourceType, object sourceData, object tableDestination, object tableName)
        {
            return InvokerService.InvokeInternal.ExecuteInt32MethodGet(this, "PivotTableWizard", sourceType, sourceData, tableDestination, tableName);
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="sourceType">optional object sourceType</param>
        /// <param name="sourceData">optional object sourceData</param>
        /// <param name="tableDestination">optional object tableDestination</param>
        /// <param name="tableName">optional object tableName</param>
        /// <param name="rowGrand">optional object rowGrand</param>
        [CustomMethod]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual Int32 PivotTableWizard(object sourceType, object sourceData, object tableDestination, object tableName, object rowGrand)
        {
            return InvokerService.InvokeInternal.ExecuteInt32MethodGet(this, "PivotTableWizard", new object[] { sourceType, sourceData, tableDestination, tableName, rowGrand });
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="sourceType">optional object sourceType</param>
        /// <param name="sourceData">optional object sourceData</param>
        /// <param name="tableDestination">optional object tableDestination</param>
        /// <param name="tableName">optional object tableName</param>
        /// <param name="rowGrand">optional object rowGrand</param>
        /// <param name="columnGrand">optional object columnGrand</param>
        [CustomMethod]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual Int32 PivotTableWizard(object sourceType, object sourceData, object tableDestination, object tableName, object rowGrand, object columnGrand)
        {
            return InvokerService.InvokeInternal.ExecuteInt32MethodGet(this, "PivotTableWizard", new object[] { sourceType, sourceData, tableDestination, tableName, rowGrand, columnGrand });
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="sourceType">optional object sourceType</param>
        /// <param name="sourceData">optional object sourceData</param>
        /// <param name="tableDestination">optional object tableDestination</param>
        /// <param name="tableName">optional object tableName</param>
        /// <param name="rowGrand">optional object rowGrand</param>
        /// <param name="columnGrand">optional object columnGrand</param>
        /// <param name="saveData">optional object saveData</param>
        [CustomMethod]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual Int32 PivotTableWizard(object sourceType, object sourceData, object tableDestination, object tableName, object rowGrand, object columnGrand, object saveData)
        {
            return InvokerService.InvokeInternal.ExecuteInt32MethodGet(this, "PivotTableWizard", new object[] { sourceType, sourceData, tableDestination, tableName, rowGrand, columnGrand, saveData });
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="sourceType">optional object sourceType</param>
        /// <param name="sourceData">optional object sourceData</param>
        /// <param name="tableDestination">optional object tableDestination</param>
        /// <param name="tableName">optional object tableName</param>
        /// <param name="rowGrand">optional object rowGrand</param>
        /// <param name="columnGrand">optional object columnGrand</param>
        /// <param name="saveData">optional object saveData</param>
        /// <param name="hasAutoFormat">optional object hasAutoFormat</param>
        [CustomMethod]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual Int32 PivotTableWizard(object sourceType, object sourceData, object tableDestination, object tableName, object rowGrand, object columnGrand, object saveData, object hasAutoFormat)
        {
            return InvokerService.InvokeInternal.ExecuteInt32MethodGet(this, "PivotTableWizard", new object[] { sourceType, sourceData, tableDestination, tableName, rowGrand, columnGrand, saveData, hasAutoFormat });
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="sourceType">optional object sourceType</param>
        /// <param name="sourceData">optional object sourceData</param>
        /// <param name="tableDestination">optional object tableDestination</param>
        /// <param name="tableName">optional object tableName</param>
        /// <param name="rowGrand">optional object rowGrand</param>
        /// <param name="columnGrand">optional object columnGrand</param>
        /// <param name="saveData">optional object saveData</param>
        /// <param name="hasAutoFormat">optional object hasAutoFormat</param>
        /// <param name="autoPage">optional object autoPage</param>
        [CustomMethod]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual Int32 PivotTableWizard(object sourceType, object sourceData, object tableDestination, object tableName, object rowGrand, object columnGrand, object saveData, object hasAutoFormat, object autoPage)
        {
            return InvokerService.InvokeInternal.ExecuteInt32MethodGet(this, "PivotTableWizard", new object[] { sourceType, sourceData, tableDestination, tableName, rowGrand, columnGrand, saveData, hasAutoFormat, autoPage });
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="sourceType">optional object sourceType</param>
        /// <param name="sourceData">optional object sourceData</param>
        /// <param name="tableDestination">optional object tableDestination</param>
        /// <param name="tableName">optional object tableName</param>
        /// <param name="rowGrand">optional object rowGrand</param>
        /// <param name="columnGrand">optional object columnGrand</param>
        /// <param name="saveData">optional object saveData</param>
        /// <param name="hasAutoFormat">optional object hasAutoFormat</param>
        /// <param name="autoPage">optional object autoPage</param>
        /// <param name="reserved">optional object reserved</param>
        [CustomMethod]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual Int32 PivotTableWizard(object sourceType, object sourceData, object tableDestination, object tableName, object rowGrand, object columnGrand, object saveData, object hasAutoFormat, object autoPage, object reserved)
        {
            return InvokerService.InvokeInternal.ExecuteInt32MethodGet(this, "PivotTableWizard", new object[] { sourceType, sourceData, tableDestination, tableName, rowGrand, columnGrand, saveData, hasAutoFormat, autoPage, reserved });
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="sourceType">optional object sourceType</param>
        /// <param name="sourceData">optional object sourceData</param>
        /// <param name="tableDestination">optional object tableDestination</param>
        /// <param name="tableName">optional object tableName</param>
        /// <param name="rowGrand">optional object rowGrand</param>
        /// <param name="columnGrand">optional object columnGrand</param>
        /// <param name="saveData">optional object saveData</param>
        /// <param name="hasAutoFormat">optional object hasAutoFormat</param>
        /// <param name="autoPage">optional object autoPage</param>
        /// <param name="reserved">optional object reserved</param>
        /// <param name="backgroundQuery">optional object backgroundQuery</param>
        [CustomMethod]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual Int32 PivotTableWizard(object sourceType, object sourceData, object tableDestination, object tableName, object rowGrand, object columnGrand, object saveData, object hasAutoFormat, object autoPage, object reserved, object backgroundQuery)
        {
            return InvokerService.InvokeInternal.ExecuteInt32MethodGet(this, "PivotTableWizard", new object[] { sourceType, sourceData, tableDestination, tableName, rowGrand, columnGrand, saveData, hasAutoFormat, autoPage, reserved, backgroundQuery });
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="sourceType">optional object sourceType</param>
        /// <param name="sourceData">optional object sourceData</param>
        /// <param name="tableDestination">optional object tableDestination</param>
        /// <param name="tableName">optional object tableName</param>
        /// <param name="rowGrand">optional object rowGrand</param>
        /// <param name="columnGrand">optional object columnGrand</param>
        /// <param name="saveData">optional object saveData</param>
        /// <param name="hasAutoFormat">optional object hasAutoFormat</param>
        /// <param name="autoPage">optional object autoPage</param>
        /// <param name="reserved">optional object reserved</param>
        /// <param name="backgroundQuery">optional object backgroundQuery</param>
        /// <param name="optimizeCache">optional object optimizeCache</param>
        [CustomMethod]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual Int32 PivotTableWizard(object sourceType, object sourceData, object tableDestination, object tableName, object rowGrand, object columnGrand, object saveData, object hasAutoFormat, object autoPage, object reserved, object backgroundQuery, object optimizeCache)
        {
            return InvokerService.InvokeInternal.ExecuteInt32MethodGet(this, "PivotTableWizard", new object[] { sourceType, sourceData, tableDestination, tableName, rowGrand, columnGrand, saveData, hasAutoFormat, autoPage, reserved, backgroundQuery, optimizeCache });
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="sourceType">optional object sourceType</param>
        /// <param name="sourceData">optional object sourceData</param>
        /// <param name="tableDestination">optional object tableDestination</param>
        /// <param name="tableName">optional object tableName</param>
        /// <param name="rowGrand">optional object rowGrand</param>
        /// <param name="columnGrand">optional object columnGrand</param>
        /// <param name="saveData">optional object saveData</param>
        /// <param name="hasAutoFormat">optional object hasAutoFormat</param>
        /// <param name="autoPage">optional object autoPage</param>
        /// <param name="reserved">optional object reserved</param>
        /// <param name="backgroundQuery">optional object backgroundQuery</param>
        /// <param name="optimizeCache">optional object optimizeCache</param>
        /// <param name="pageFieldOrder">optional object pageFieldOrder</param>
        [CustomMethod]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual Int32 PivotTableWizard(object sourceType, object sourceData, object tableDestination, object tableName, object rowGrand, object columnGrand, object saveData, object hasAutoFormat, object autoPage, object reserved, object backgroundQuery, object optimizeCache, object pageFieldOrder)
        {
            return InvokerService.InvokeInternal.ExecuteInt32MethodGet(this, "PivotTableWizard", new object[] { sourceType, sourceData, tableDestination, tableName, rowGrand, columnGrand, saveData, hasAutoFormat, autoPage, reserved, backgroundQuery, optimizeCache, pageFieldOrder });
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="sourceType">optional object sourceType</param>
        /// <param name="sourceData">optional object sourceData</param>
        /// <param name="tableDestination">optional object tableDestination</param>
        /// <param name="tableName">optional object tableName</param>
        /// <param name="rowGrand">optional object rowGrand</param>
        /// <param name="columnGrand">optional object columnGrand</param>
        /// <param name="saveData">optional object saveData</param>
        /// <param name="hasAutoFormat">optional object hasAutoFormat</param>
        /// <param name="autoPage">optional object autoPage</param>
        /// <param name="reserved">optional object reserved</param>
        /// <param name="backgroundQuery">optional object backgroundQuery</param>
        /// <param name="optimizeCache">optional object optimizeCache</param>
        /// <param name="pageFieldOrder">optional object pageFieldOrder</param>
        /// <param name="pageFieldWrapCount">optional object pageFieldWrapCount</param>
        [CustomMethod]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual Int32 PivotTableWizard(object sourceType, object sourceData, object tableDestination, object tableName, object rowGrand, object columnGrand, object saveData, object hasAutoFormat, object autoPage, object reserved, object backgroundQuery, object optimizeCache, object pageFieldOrder, object pageFieldWrapCount)
        {
            return InvokerService.InvokeInternal.ExecuteInt32MethodGet(this, "PivotTableWizard", new object[] { sourceType, sourceData, tableDestination, tableName, rowGrand, columnGrand, saveData, hasAutoFormat, autoPage, reserved, backgroundQuery, optimizeCache, pageFieldOrder, pageFieldWrapCount });
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="sourceType">optional object sourceType</param>
        /// <param name="sourceData">optional object sourceData</param>
        /// <param name="tableDestination">optional object tableDestination</param>
        /// <param name="tableName">optional object tableName</param>
        /// <param name="rowGrand">optional object rowGrand</param>
        /// <param name="columnGrand">optional object columnGrand</param>
        /// <param name="saveData">optional object saveData</param>
        /// <param name="hasAutoFormat">optional object hasAutoFormat</param>
        /// <param name="autoPage">optional object autoPage</param>
        /// <param name="reserved">optional object reserved</param>
        /// <param name="backgroundQuery">optional object backgroundQuery</param>
        /// <param name="optimizeCache">optional object optimizeCache</param>
        /// <param name="pageFieldOrder">optional object pageFieldOrder</param>
        /// <param name="pageFieldWrapCount">optional object pageFieldWrapCount</param>
        /// <param name="readData">optional object readData</param>
        [CustomMethod]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual Int32 PivotTableWizard(object sourceType, object sourceData, object tableDestination, object tableName, object rowGrand, object columnGrand, object saveData, object hasAutoFormat, object autoPage, object reserved, object backgroundQuery, object optimizeCache, object pageFieldOrder, object pageFieldWrapCount, object readData)
        {
            return InvokerService.InvokeInternal.ExecuteInt32MethodGet(this, "PivotTableWizard", new object[] { sourceType, sourceData, tableDestination, tableName, rowGrand, columnGrand, saveData, hasAutoFormat, autoPage, reserved, backgroundQuery, optimizeCache, pageFieldOrder, pageFieldWrapCount, readData });
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="name">string name</param>
        /// <param name="mode">optional NetOffice.ExcelApi.Enums.XlPTSelectionMode Mode = 0</param>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual Int32 PivotSelect(string name, object mode)
        {
            return InvokerService.InvokeInternal.ExecuteInt32MethodGet(this, "PivotSelect", name, mode);
        }

        /// <summary>
        /// SupportByVersion Excel 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="name">string name</param>
        /// <param name="mode">optional NetOffice.ExcelApi.Enums.XlPTSelectionMode Mode = 0</param>
        /// <param name="useStandardName">optional object useStandardName</param>
        [SupportByVersion("Excel", 10, 11, 12, 14, 15, 16)]
        public virtual Int32 PivotSelect(string name, object mode, object useStandardName)
        {
            return InvokerService.InvokeInternal.ExecuteInt32MethodGet(this, "PivotSelect", name, mode, useStandardName);
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="name">string name</param>
        [CustomMethod]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual Int32 PivotSelect(string name)
        {
            return InvokerService.InvokeInternal.ExecuteInt32MethodGet(this, "PivotSelect", name);
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual Int32 Update()
        {
            return InvokerService.InvokeInternal.ExecuteInt32MethodGet(this, "Update");
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="format">NetOffice.ExcelApi.Enums.xlPivotFormatType format</param>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual Int32 Format(NetOffice.ExcelApi.Enums.xlPivotFormatType format)
        {
            return InvokerService.InvokeInternal.ExecuteInt32MethodGet(this, "Format", format);
        }

        /// <summary>
        /// SupportByVersion Excel 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="name">string name</param>
        /// <param name="mode">optional NetOffice.ExcelApi.Enums.XlPTSelectionMode Mode = 0</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [SupportByVersion("Excel", 10, 11, 12, 14, 15, 16)]
        public virtual Int32 _PivotSelect(string name, object mode)
        {
            return InvokerService.InvokeInternal.ExecuteInt32MethodGet(this, "_PivotSelect", name, mode);
        }

        /// <summary>
        /// SupportByVersion Excel 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="name">string name</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [CustomMethod]
        [SupportByVersion("Excel", 10, 11, 12, 14, 15, 16)]
        public virtual Int32 _PivotSelect(string name)
        {
            return InvokerService.InvokeInternal.ExecuteInt32MethodGet(this, "_PivotSelect", name);
        }

        /// <summary>
        /// SupportByVersion Excel 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="dataField">optional object dataField</param>
        /// <param name="field1">optional object field1</param>
        /// <param name="item1">optional object item1</param>
        /// <param name="field2">optional object field2</param>
        /// <param name="item2">optional object item2</param>
        /// <param name="field3">optional object field3</param>
        /// <param name="item3">optional object item3</param>
        /// <param name="field4">optional object field4</param>
        /// <param name="item4">optional object item4</param>
        /// <param name="field5">optional object field5</param>
        /// <param name="item5">optional object item5</param>
        /// <param name="field6">optional object field6</param>
        /// <param name="item6">optional object item6</param>
        /// <param name="field7">optional object field7</param>
        /// <param name="item7">optional object item7</param>
        /// <param name="field8">optional object field8</param>
        /// <param name="item8">optional object item8</param>
        /// <param name="field9">optional object field9</param>
        /// <param name="item9">optional object item9</param>
        /// <param name="field10">optional object field10</param>
        /// <param name="item10">optional object item10</param>
        /// <param name="field11">optional object field11</param>
        /// <param name="item11">optional object item11</param>
        /// <param name="field12">optional object field12</param>
        /// <param name="item12">optional object item12</param>
        /// <param name="field13">optional object field13</param>
        /// <param name="item13">optional object item13</param>
        /// <param name="field14">optional object field14</param>
        /// <param name="item14">optional object item14</param>
        [SupportByVersion("Excel", 10, 11, 12, 14, 15, 16)]
        public virtual NetOffice.ExcelApi.Range GetPivotData(object dataField, object field1, object item1, object field2, object item2, object field3, object item3, object field4, object item4, object field5, object item5, object field6, object item6, object field7, object item7, object field8, object item8, object field9, object item9, object field10, object item10, object field11, object item11, object field12, object item12, object field13, object item13, object field14, object item14)
        {
            return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.ExcelApi.Range>(this, "GetPivotData", typeof(NetOffice.ExcelApi.Range), new object[] { dataField, field1, item1, field2, item2, field3, item3, field4, item4, field5, item5, field6, item6, field7, item7, field8, item8, field9, item9, field10, item10, field11, item11, field12, item12, field13, item13, field14, item14 });
        }

        /// <summary>
        /// SupportByVersion Excel 10, 11, 12, 14, 15, 16
        /// </summary>
        [CustomMethod]
        [SupportByVersion("Excel", 10, 11, 12, 14, 15, 16)]
        public virtual NetOffice.ExcelApi.Range GetPivotData()
        {
            return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.ExcelApi.Range>(this, "GetPivotData", typeof(NetOffice.ExcelApi.Range));
        }

        /// <summary>
        /// SupportByVersion Excel 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="dataField">optional object dataField</param>
        [CustomMethod]
        [SupportByVersion("Excel", 10, 11, 12, 14, 15, 16)]
        public virtual NetOffice.ExcelApi.Range GetPivotData(object dataField)
        {
            return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.ExcelApi.Range>(this, "GetPivotData", typeof(NetOffice.ExcelApi.Range), dataField);
        }

        /// <summary>
        /// SupportByVersion Excel 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="dataField">optional object dataField</param>
        /// <param name="field1">optional object field1</param>
        [CustomMethod]
        [SupportByVersion("Excel", 10, 11, 12, 14, 15, 16)]
        public virtual NetOffice.ExcelApi.Range GetPivotData(object dataField, object field1)
        {
            return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.ExcelApi.Range>(this, "GetPivotData", typeof(NetOffice.ExcelApi.Range), dataField, field1);
        }

        /// <summary>
        /// SupportByVersion Excel 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="dataField">optional object dataField</param>
        /// <param name="field1">optional object field1</param>
        /// <param name="item1">optional object item1</param>
        [CustomMethod]
        [SupportByVersion("Excel", 10, 11, 12, 14, 15, 16)]
        public virtual NetOffice.ExcelApi.Range GetPivotData(object dataField, object field1, object item1)
        {
            return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.ExcelApi.Range>(this, "GetPivotData", typeof(NetOffice.ExcelApi.Range), dataField, field1, item1);
        }

        /// <summary>
        /// SupportByVersion Excel 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="dataField">optional object dataField</param>
        /// <param name="field1">optional object field1</param>
        /// <param name="item1">optional object item1</param>
        /// <param name="field2">optional object field2</param>
        [CustomMethod]
        [SupportByVersion("Excel", 10, 11, 12, 14, 15, 16)]
        public virtual NetOffice.ExcelApi.Range GetPivotData(object dataField, object field1, object item1, object field2)
        {
            return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.ExcelApi.Range>(this, "GetPivotData", typeof(NetOffice.ExcelApi.Range), dataField, field1, item1, field2);
        }

        /// <summary>
        /// SupportByVersion Excel 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="dataField">optional object dataField</param>
        /// <param name="field1">optional object field1</param>
        /// <param name="item1">optional object item1</param>
        /// <param name="field2">optional object field2</param>
        /// <param name="item2">optional object item2</param>
        [CustomMethod]
        [SupportByVersion("Excel", 10, 11, 12, 14, 15, 16)]
        public virtual NetOffice.ExcelApi.Range GetPivotData(object dataField, object field1, object item1, object field2, object item2)
        {
            return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.ExcelApi.Range>(this, "GetPivotData", typeof(NetOffice.ExcelApi.Range), new object[] { dataField, field1, item1, field2, item2 });
        }

        /// <summary>
        /// SupportByVersion Excel 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="dataField">optional object dataField</param>
        /// <param name="field1">optional object field1</param>
        /// <param name="item1">optional object item1</param>
        /// <param name="field2">optional object field2</param>
        /// <param name="item2">optional object item2</param>
        /// <param name="field3">optional object field3</param>
        [CustomMethod]
        [SupportByVersion("Excel", 10, 11, 12, 14, 15, 16)]
        public virtual NetOffice.ExcelApi.Range GetPivotData(object dataField, object field1, object item1, object field2, object item2, object field3)
        {
            return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.ExcelApi.Range>(this, "GetPivotData", typeof(NetOffice.ExcelApi.Range), new object[] { dataField, field1, item1, field2, item2, field3 });
        }

        /// <summary>
        /// SupportByVersion Excel 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="dataField">optional object dataField</param>
        /// <param name="field1">optional object field1</param>
        /// <param name="item1">optional object item1</param>
        /// <param name="field2">optional object field2</param>
        /// <param name="item2">optional object item2</param>
        /// <param name="field3">optional object field3</param>
        /// <param name="item3">optional object item3</param>
        [CustomMethod]
        [SupportByVersion("Excel", 10, 11, 12, 14, 15, 16)]
        public virtual NetOffice.ExcelApi.Range GetPivotData(object dataField, object field1, object item1, object field2, object item2, object field3, object item3)
        {
            return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.ExcelApi.Range>(this, "GetPivotData", typeof(NetOffice.ExcelApi.Range), new object[] { dataField, field1, item1, field2, item2, field3, item3 });
        }

        /// <summary>
        /// SupportByVersion Excel 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="dataField">optional object dataField</param>
        /// <param name="field1">optional object field1</param>
        /// <param name="item1">optional object item1</param>
        /// <param name="field2">optional object field2</param>
        /// <param name="item2">optional object item2</param>
        /// <param name="field3">optional object field3</param>
        /// <param name="item3">optional object item3</param>
        /// <param name="field4">optional object field4</param>
        [CustomMethod]
        [SupportByVersion("Excel", 10, 11, 12, 14, 15, 16)]
        public virtual NetOffice.ExcelApi.Range GetPivotData(object dataField, object field1, object item1, object field2, object item2, object field3, object item3, object field4)
        {
            return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.ExcelApi.Range>(this, "GetPivotData", typeof(NetOffice.ExcelApi.Range), new object[] { dataField, field1, item1, field2, item2, field3, item3, field4 });
        }

        /// <summary>
        /// SupportByVersion Excel 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="dataField">optional object dataField</param>
        /// <param name="field1">optional object field1</param>
        /// <param name="item1">optional object item1</param>
        /// <param name="field2">optional object field2</param>
        /// <param name="item2">optional object item2</param>
        /// <param name="field3">optional object field3</param>
        /// <param name="item3">optional object item3</param>
        /// <param name="field4">optional object field4</param>
        /// <param name="item4">optional object item4</param>
        [CustomMethod]
        [SupportByVersion("Excel", 10, 11, 12, 14, 15, 16)]
        public virtual NetOffice.ExcelApi.Range GetPivotData(object dataField, object field1, object item1, object field2, object item2, object field3, object item3, object field4, object item4)
        {
            return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.ExcelApi.Range>(this, "GetPivotData", typeof(NetOffice.ExcelApi.Range), new object[] { dataField, field1, item1, field2, item2, field3, item3, field4, item4 });
        }

        /// <summary>
        /// SupportByVersion Excel 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="dataField">optional object dataField</param>
        /// <param name="field1">optional object field1</param>
        /// <param name="item1">optional object item1</param>
        /// <param name="field2">optional object field2</param>
        /// <param name="item2">optional object item2</param>
        /// <param name="field3">optional object field3</param>
        /// <param name="item3">optional object item3</param>
        /// <param name="field4">optional object field4</param>
        /// <param name="item4">optional object item4</param>
        /// <param name="field5">optional object field5</param>
        [CustomMethod]
        [SupportByVersion("Excel", 10, 11, 12, 14, 15, 16)]
        public virtual NetOffice.ExcelApi.Range GetPivotData(object dataField, object field1, object item1, object field2, object item2, object field3, object item3, object field4, object item4, object field5)
        {
            return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.ExcelApi.Range>(this, "GetPivotData", typeof(NetOffice.ExcelApi.Range), new object[] { dataField, field1, item1, field2, item2, field3, item3, field4, item4, field5 });
        }

        /// <summary>
        /// SupportByVersion Excel 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="dataField">optional object dataField</param>
        /// <param name="field1">optional object field1</param>
        /// <param name="item1">optional object item1</param>
        /// <param name="field2">optional object field2</param>
        /// <param name="item2">optional object item2</param>
        /// <param name="field3">optional object field3</param>
        /// <param name="item3">optional object item3</param>
        /// <param name="field4">optional object field4</param>
        /// <param name="item4">optional object item4</param>
        /// <param name="field5">optional object field5</param>
        /// <param name="item5">optional object item5</param>
        [CustomMethod]
        [SupportByVersion("Excel", 10, 11, 12, 14, 15, 16)]
        public virtual NetOffice.ExcelApi.Range GetPivotData(object dataField, object field1, object item1, object field2, object item2, object field3, object item3, object field4, object item4, object field5, object item5)
        {
            return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.ExcelApi.Range>(this, "GetPivotData", typeof(NetOffice.ExcelApi.Range), new object[] { dataField, field1, item1, field2, item2, field3, item3, field4, item4, field5, item5 });
        }

        /// <summary>
        /// SupportByVersion Excel 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="dataField">optional object dataField</param>
        /// <param name="field1">optional object field1</param>
        /// <param name="item1">optional object item1</param>
        /// <param name="field2">optional object field2</param>
        /// <param name="item2">optional object item2</param>
        /// <param name="field3">optional object field3</param>
        /// <param name="item3">optional object item3</param>
        /// <param name="field4">optional object field4</param>
        /// <param name="item4">optional object item4</param>
        /// <param name="field5">optional object field5</param>
        /// <param name="item5">optional object item5</param>
        /// <param name="field6">optional object field6</param>
        [CustomMethod]
        [SupportByVersion("Excel", 10, 11, 12, 14, 15, 16)]
        public virtual NetOffice.ExcelApi.Range GetPivotData(object dataField, object field1, object item1, object field2, object item2, object field3, object item3, object field4, object item4, object field5, object item5, object field6)
        {
            return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.ExcelApi.Range>(this, "GetPivotData", typeof(NetOffice.ExcelApi.Range), new object[] { dataField, field1, item1, field2, item2, field3, item3, field4, item4, field5, item5, field6 });
        }

        /// <summary>
        /// SupportByVersion Excel 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="dataField">optional object dataField</param>
        /// <param name="field1">optional object field1</param>
        /// <param name="item1">optional object item1</param>
        /// <param name="field2">optional object field2</param>
        /// <param name="item2">optional object item2</param>
        /// <param name="field3">optional object field3</param>
        /// <param name="item3">optional object item3</param>
        /// <param name="field4">optional object field4</param>
        /// <param name="item4">optional object item4</param>
        /// <param name="field5">optional object field5</param>
        /// <param name="item5">optional object item5</param>
        /// <param name="field6">optional object field6</param>
        /// <param name="item6">optional object item6</param>
        [CustomMethod]
        [SupportByVersion("Excel", 10, 11, 12, 14, 15, 16)]
        public virtual NetOffice.ExcelApi.Range GetPivotData(object dataField, object field1, object item1, object field2, object item2, object field3, object item3, object field4, object item4, object field5, object item5, object field6, object item6)
        {
            return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.ExcelApi.Range>(this, "GetPivotData", typeof(NetOffice.ExcelApi.Range), new object[] { dataField, field1, item1, field2, item2, field3, item3, field4, item4, field5, item5, field6, item6 });
        }

        /// <summary>
        /// SupportByVersion Excel 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="dataField">optional object dataField</param>
        /// <param name="field1">optional object field1</param>
        /// <param name="item1">optional object item1</param>
        /// <param name="field2">optional object field2</param>
        /// <param name="item2">optional object item2</param>
        /// <param name="field3">optional object field3</param>
        /// <param name="item3">optional object item3</param>
        /// <param name="field4">optional object field4</param>
        /// <param name="item4">optional object item4</param>
        /// <param name="field5">optional object field5</param>
        /// <param name="item5">optional object item5</param>
        /// <param name="field6">optional object field6</param>
        /// <param name="item6">optional object item6</param>
        /// <param name="field7">optional object field7</param>
        [CustomMethod]
        [SupportByVersion("Excel", 10, 11, 12, 14, 15, 16)]
        public virtual NetOffice.ExcelApi.Range GetPivotData(object dataField, object field1, object item1, object field2, object item2, object field3, object item3, object field4, object item4, object field5, object item5, object field6, object item6, object field7)
        {
            return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.ExcelApi.Range>(this, "GetPivotData", typeof(NetOffice.ExcelApi.Range), new object[] { dataField, field1, item1, field2, item2, field3, item3, field4, item4, field5, item5, field6, item6, field7 });
        }

        /// <summary>
        /// SupportByVersion Excel 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="dataField">optional object dataField</param>
        /// <param name="field1">optional object field1</param>
        /// <param name="item1">optional object item1</param>
        /// <param name="field2">optional object field2</param>
        /// <param name="item2">optional object item2</param>
        /// <param name="field3">optional object field3</param>
        /// <param name="item3">optional object item3</param>
        /// <param name="field4">optional object field4</param>
        /// <param name="item4">optional object item4</param>
        /// <param name="field5">optional object field5</param>
        /// <param name="item5">optional object item5</param>
        /// <param name="field6">optional object field6</param>
        /// <param name="item6">optional object item6</param>
        /// <param name="field7">optional object field7</param>
        /// <param name="item7">optional object item7</param>
        [CustomMethod]
        [SupportByVersion("Excel", 10, 11, 12, 14, 15, 16)]
        public virtual NetOffice.ExcelApi.Range GetPivotData(object dataField, object field1, object item1, object field2, object item2, object field3, object item3, object field4, object item4, object field5, object item5, object field6, object item6, object field7, object item7)
        {
            return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.ExcelApi.Range>(this, "GetPivotData", typeof(NetOffice.ExcelApi.Range), new object[] { dataField, field1, item1, field2, item2, field3, item3, field4, item4, field5, item5, field6, item6, field7, item7 });
        }

        /// <summary>
        /// SupportByVersion Excel 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="dataField">optional object dataField</param>
        /// <param name="field1">optional object field1</param>
        /// <param name="item1">optional object item1</param>
        /// <param name="field2">optional object field2</param>
        /// <param name="item2">optional object item2</param>
        /// <param name="field3">optional object field3</param>
        /// <param name="item3">optional object item3</param>
        /// <param name="field4">optional object field4</param>
        /// <param name="item4">optional object item4</param>
        /// <param name="field5">optional object field5</param>
        /// <param name="item5">optional object item5</param>
        /// <param name="field6">optional object field6</param>
        /// <param name="item6">optional object item6</param>
        /// <param name="field7">optional object field7</param>
        /// <param name="item7">optional object item7</param>
        /// <param name="field8">optional object field8</param>
        [CustomMethod]
        [SupportByVersion("Excel", 10, 11, 12, 14, 15, 16)]
        public virtual NetOffice.ExcelApi.Range GetPivotData(object dataField, object field1, object item1, object field2, object item2, object field3, object item3, object field4, object item4, object field5, object item5, object field6, object item6, object field7, object item7, object field8)
        {
            return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.ExcelApi.Range>(this, "GetPivotData", typeof(NetOffice.ExcelApi.Range), new object[] { dataField, field1, item1, field2, item2, field3, item3, field4, item4, field5, item5, field6, item6, field7, item7, field8 });
        }

        /// <summary>
        /// SupportByVersion Excel 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="dataField">optional object dataField</param>
        /// <param name="field1">optional object field1</param>
        /// <param name="item1">optional object item1</param>
        /// <param name="field2">optional object field2</param>
        /// <param name="item2">optional object item2</param>
        /// <param name="field3">optional object field3</param>
        /// <param name="item3">optional object item3</param>
        /// <param name="field4">optional object field4</param>
        /// <param name="item4">optional object item4</param>
        /// <param name="field5">optional object field5</param>
        /// <param name="item5">optional object item5</param>
        /// <param name="field6">optional object field6</param>
        /// <param name="item6">optional object item6</param>
        /// <param name="field7">optional object field7</param>
        /// <param name="item7">optional object item7</param>
        /// <param name="field8">optional object field8</param>
        /// <param name="item8">optional object item8</param>
        [CustomMethod]
        [SupportByVersion("Excel", 10, 11, 12, 14, 15, 16)]
        public virtual NetOffice.ExcelApi.Range GetPivotData(object dataField, object field1, object item1, object field2, object item2, object field3, object item3, object field4, object item4, object field5, object item5, object field6, object item6, object field7, object item7, object field8, object item8)
        {
            return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.ExcelApi.Range>(this, "GetPivotData", typeof(NetOffice.ExcelApi.Range), new object[] { dataField, field1, item1, field2, item2, field3, item3, field4, item4, field5, item5, field6, item6, field7, item7, field8, item8 });
        }

        /// <summary>
        /// SupportByVersion Excel 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="dataField">optional object dataField</param>
        /// <param name="field1">optional object field1</param>
        /// <param name="item1">optional object item1</param>
        /// <param name="field2">optional object field2</param>
        /// <param name="item2">optional object item2</param>
        /// <param name="field3">optional object field3</param>
        /// <param name="item3">optional object item3</param>
        /// <param name="field4">optional object field4</param>
        /// <param name="item4">optional object item4</param>
        /// <param name="field5">optional object field5</param>
        /// <param name="item5">optional object item5</param>
        /// <param name="field6">optional object field6</param>
        /// <param name="item6">optional object item6</param>
        /// <param name="field7">optional object field7</param>
        /// <param name="item7">optional object item7</param>
        /// <param name="field8">optional object field8</param>
        /// <param name="item8">optional object item8</param>
        /// <param name="field9">optional object field9</param>
        [CustomMethod]
        [SupportByVersion("Excel", 10, 11, 12, 14, 15, 16)]
        public virtual NetOffice.ExcelApi.Range GetPivotData(object dataField, object field1, object item1, object field2, object item2, object field3, object item3, object field4, object item4, object field5, object item5, object field6, object item6, object field7, object item7, object field8, object item8, object field9)
        {
            return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.ExcelApi.Range>(this, "GetPivotData", typeof(NetOffice.ExcelApi.Range), new object[] { dataField, field1, item1, field2, item2, field3, item3, field4, item4, field5, item5, field6, item6, field7, item7, field8, item8, field9 });
        }

        /// <summary>
        /// SupportByVersion Excel 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="dataField">optional object dataField</param>
        /// <param name="field1">optional object field1</param>
        /// <param name="item1">optional object item1</param>
        /// <param name="field2">optional object field2</param>
        /// <param name="item2">optional object item2</param>
        /// <param name="field3">optional object field3</param>
        /// <param name="item3">optional object item3</param>
        /// <param name="field4">optional object field4</param>
        /// <param name="item4">optional object item4</param>
        /// <param name="field5">optional object field5</param>
        /// <param name="item5">optional object item5</param>
        /// <param name="field6">optional object field6</param>
        /// <param name="item6">optional object item6</param>
        /// <param name="field7">optional object field7</param>
        /// <param name="item7">optional object item7</param>
        /// <param name="field8">optional object field8</param>
        /// <param name="item8">optional object item8</param>
        /// <param name="field9">optional object field9</param>
        /// <param name="item9">optional object item9</param>
        [CustomMethod]
        [SupportByVersion("Excel", 10, 11, 12, 14, 15, 16)]
        public virtual NetOffice.ExcelApi.Range GetPivotData(object dataField, object field1, object item1, object field2, object item2, object field3, object item3, object field4, object item4, object field5, object item5, object field6, object item6, object field7, object item7, object field8, object item8, object field9, object item9)
        {
            return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.ExcelApi.Range>(this, "GetPivotData", typeof(NetOffice.ExcelApi.Range), new object[] { dataField, field1, item1, field2, item2, field3, item3, field4, item4, field5, item5, field6, item6, field7, item7, field8, item8, field9, item9 });
        }

        /// <summary>
        /// SupportByVersion Excel 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="dataField">optional object dataField</param>
        /// <param name="field1">optional object field1</param>
        /// <param name="item1">optional object item1</param>
        /// <param name="field2">optional object field2</param>
        /// <param name="item2">optional object item2</param>
        /// <param name="field3">optional object field3</param>
        /// <param name="item3">optional object item3</param>
        /// <param name="field4">optional object field4</param>
        /// <param name="item4">optional object item4</param>
        /// <param name="field5">optional object field5</param>
        /// <param name="item5">optional object item5</param>
        /// <param name="field6">optional object field6</param>
        /// <param name="item6">optional object item6</param>
        /// <param name="field7">optional object field7</param>
        /// <param name="item7">optional object item7</param>
        /// <param name="field8">optional object field8</param>
        /// <param name="item8">optional object item8</param>
        /// <param name="field9">optional object field9</param>
        /// <param name="item9">optional object item9</param>
        /// <param name="field10">optional object field10</param>
        [CustomMethod]
        [SupportByVersion("Excel", 10, 11, 12, 14, 15, 16)]
        public virtual NetOffice.ExcelApi.Range GetPivotData(object dataField, object field1, object item1, object field2, object item2, object field3, object item3, object field4, object item4, object field5, object item5, object field6, object item6, object field7, object item7, object field8, object item8, object field9, object item9, object field10)
        {
            return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.ExcelApi.Range>(this, "GetPivotData", typeof(NetOffice.ExcelApi.Range), new object[] { dataField, field1, item1, field2, item2, field3, item3, field4, item4, field5, item5, field6, item6, field7, item7, field8, item8, field9, item9, field10 });
        }

        /// <summary>
        /// SupportByVersion Excel 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="dataField">optional object dataField</param>
        /// <param name="field1">optional object field1</param>
        /// <param name="item1">optional object item1</param>
        /// <param name="field2">optional object field2</param>
        /// <param name="item2">optional object item2</param>
        /// <param name="field3">optional object field3</param>
        /// <param name="item3">optional object item3</param>
        /// <param name="field4">optional object field4</param>
        /// <param name="item4">optional object item4</param>
        /// <param name="field5">optional object field5</param>
        /// <param name="item5">optional object item5</param>
        /// <param name="field6">optional object field6</param>
        /// <param name="item6">optional object item6</param>
        /// <param name="field7">optional object field7</param>
        /// <param name="item7">optional object item7</param>
        /// <param name="field8">optional object field8</param>
        /// <param name="item8">optional object item8</param>
        /// <param name="field9">optional object field9</param>
        /// <param name="item9">optional object item9</param>
        /// <param name="field10">optional object field10</param>
        /// <param name="item10">optional object item10</param>
        [CustomMethod]
        [SupportByVersion("Excel", 10, 11, 12, 14, 15, 16)]
        public virtual NetOffice.ExcelApi.Range GetPivotData(object dataField, object field1, object item1, object field2, object item2, object field3, object item3, object field4, object item4, object field5, object item5, object field6, object item6, object field7, object item7, object field8, object item8, object field9, object item9, object field10, object item10)
        {
            return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.ExcelApi.Range>(this, "GetPivotData", typeof(NetOffice.ExcelApi.Range), new object[] { dataField, field1, item1, field2, item2, field3, item3, field4, item4, field5, item5, field6, item6, field7, item7, field8, item8, field9, item9, field10, item10 });
        }

        /// <summary>
        /// SupportByVersion Excel 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="dataField">optional object dataField</param>
        /// <param name="field1">optional object field1</param>
        /// <param name="item1">optional object item1</param>
        /// <param name="field2">optional object field2</param>
        /// <param name="item2">optional object item2</param>
        /// <param name="field3">optional object field3</param>
        /// <param name="item3">optional object item3</param>
        /// <param name="field4">optional object field4</param>
        /// <param name="item4">optional object item4</param>
        /// <param name="field5">optional object field5</param>
        /// <param name="item5">optional object item5</param>
        /// <param name="field6">optional object field6</param>
        /// <param name="item6">optional object item6</param>
        /// <param name="field7">optional object field7</param>
        /// <param name="item7">optional object item7</param>
        /// <param name="field8">optional object field8</param>
        /// <param name="item8">optional object item8</param>
        /// <param name="field9">optional object field9</param>
        /// <param name="item9">optional object item9</param>
        /// <param name="field10">optional object field10</param>
        /// <param name="item10">optional object item10</param>
        /// <param name="field11">optional object field11</param>
        [CustomMethod]
        [SupportByVersion("Excel", 10, 11, 12, 14, 15, 16)]
        public virtual NetOffice.ExcelApi.Range GetPivotData(object dataField, object field1, object item1, object field2, object item2, object field3, object item3, object field4, object item4, object field5, object item5, object field6, object item6, object field7, object item7, object field8, object item8, object field9, object item9, object field10, object item10, object field11)
        {
            return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.ExcelApi.Range>(this, "GetPivotData", typeof(NetOffice.ExcelApi.Range), new object[] { dataField, field1, item1, field2, item2, field3, item3, field4, item4, field5, item5, field6, item6, field7, item7, field8, item8, field9, item9, field10, item10, field11 });
        }

        /// <summary>
        /// SupportByVersion Excel 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="dataField">optional object dataField</param>
        /// <param name="field1">optional object field1</param>
        /// <param name="item1">optional object item1</param>
        /// <param name="field2">optional object field2</param>
        /// <param name="item2">optional object item2</param>
        /// <param name="field3">optional object field3</param>
        /// <param name="item3">optional object item3</param>
        /// <param name="field4">optional object field4</param>
        /// <param name="item4">optional object item4</param>
        /// <param name="field5">optional object field5</param>
        /// <param name="item5">optional object item5</param>
        /// <param name="field6">optional object field6</param>
        /// <param name="item6">optional object item6</param>
        /// <param name="field7">optional object field7</param>
        /// <param name="item7">optional object item7</param>
        /// <param name="field8">optional object field8</param>
        /// <param name="item8">optional object item8</param>
        /// <param name="field9">optional object field9</param>
        /// <param name="item9">optional object item9</param>
        /// <param name="field10">optional object field10</param>
        /// <param name="item10">optional object item10</param>
        /// <param name="field11">optional object field11</param>
        /// <param name="item11">optional object item11</param>
        [CustomMethod]
        [SupportByVersion("Excel", 10, 11, 12, 14, 15, 16)]
        public virtual NetOffice.ExcelApi.Range GetPivotData(object dataField, object field1, object item1, object field2, object item2, object field3, object item3, object field4, object item4, object field5, object item5, object field6, object item6, object field7, object item7, object field8, object item8, object field9, object item9, object field10, object item10, object field11, object item11)
        {
            return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.ExcelApi.Range>(this, "GetPivotData", typeof(NetOffice.ExcelApi.Range), new object[] { dataField, field1, item1, field2, item2, field3, item3, field4, item4, field5, item5, field6, item6, field7, item7, field8, item8, field9, item9, field10, item10, field11, item11 });
        }

        /// <summary>
        /// SupportByVersion Excel 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="dataField">optional object dataField</param>
        /// <param name="field1">optional object field1</param>
        /// <param name="item1">optional object item1</param>
        /// <param name="field2">optional object field2</param>
        /// <param name="item2">optional object item2</param>
        /// <param name="field3">optional object field3</param>
        /// <param name="item3">optional object item3</param>
        /// <param name="field4">optional object field4</param>
        /// <param name="item4">optional object item4</param>
        /// <param name="field5">optional object field5</param>
        /// <param name="item5">optional object item5</param>
        /// <param name="field6">optional object field6</param>
        /// <param name="item6">optional object item6</param>
        /// <param name="field7">optional object field7</param>
        /// <param name="item7">optional object item7</param>
        /// <param name="field8">optional object field8</param>
        /// <param name="item8">optional object item8</param>
        /// <param name="field9">optional object field9</param>
        /// <param name="item9">optional object item9</param>
        /// <param name="field10">optional object field10</param>
        /// <param name="item10">optional object item10</param>
        /// <param name="field11">optional object field11</param>
        /// <param name="item11">optional object item11</param>
        /// <param name="field12">optional object field12</param>
        [CustomMethod]
        [SupportByVersion("Excel", 10, 11, 12, 14, 15, 16)]
        public virtual NetOffice.ExcelApi.Range GetPivotData(object dataField, object field1, object item1, object field2, object item2, object field3, object item3, object field4, object item4, object field5, object item5, object field6, object item6, object field7, object item7, object field8, object item8, object field9, object item9, object field10, object item10, object field11, object item11, object field12)
        {
            return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.ExcelApi.Range>(this, "GetPivotData", typeof(NetOffice.ExcelApi.Range), new object[] { dataField, field1, item1, field2, item2, field3, item3, field4, item4, field5, item5, field6, item6, field7, item7, field8, item8, field9, item9, field10, item10, field11, item11, field12 });
        }

        /// <summary>
        /// SupportByVersion Excel 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="dataField">optional object dataField</param>
        /// <param name="field1">optional object field1</param>
        /// <param name="item1">optional object item1</param>
        /// <param name="field2">optional object field2</param>
        /// <param name="item2">optional object item2</param>
        /// <param name="field3">optional object field3</param>
        /// <param name="item3">optional object item3</param>
        /// <param name="field4">optional object field4</param>
        /// <param name="item4">optional object item4</param>
        /// <param name="field5">optional object field5</param>
        /// <param name="item5">optional object item5</param>
        /// <param name="field6">optional object field6</param>
        /// <param name="item6">optional object item6</param>
        /// <param name="field7">optional object field7</param>
        /// <param name="item7">optional object item7</param>
        /// <param name="field8">optional object field8</param>
        /// <param name="item8">optional object item8</param>
        /// <param name="field9">optional object field9</param>
        /// <param name="item9">optional object item9</param>
        /// <param name="field10">optional object field10</param>
        /// <param name="item10">optional object item10</param>
        /// <param name="field11">optional object field11</param>
        /// <param name="item11">optional object item11</param>
        /// <param name="field12">optional object field12</param>
        /// <param name="item12">optional object item12</param>
        [CustomMethod]
        [SupportByVersion("Excel", 10, 11, 12, 14, 15, 16)]
        public virtual NetOffice.ExcelApi.Range GetPivotData(object dataField, object field1, object item1, object field2, object item2, object field3, object item3, object field4, object item4, object field5, object item5, object field6, object item6, object field7, object item7, object field8, object item8, object field9, object item9, object field10, object item10, object field11, object item11, object field12, object item12)
        {
            return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.ExcelApi.Range>(this, "GetPivotData", typeof(NetOffice.ExcelApi.Range), new object[] { dataField, field1, item1, field2, item2, field3, item3, field4, item4, field5, item5, field6, item6, field7, item7, field8, item8, field9, item9, field10, item10, field11, item11, field12, item12 });
        }

        /// <summary>
        /// SupportByVersion Excel 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="dataField">optional object dataField</param>
        /// <param name="field1">optional object field1</param>
        /// <param name="item1">optional object item1</param>
        /// <param name="field2">optional object field2</param>
        /// <param name="item2">optional object item2</param>
        /// <param name="field3">optional object field3</param>
        /// <param name="item3">optional object item3</param>
        /// <param name="field4">optional object field4</param>
        /// <param name="item4">optional object item4</param>
        /// <param name="field5">optional object field5</param>
        /// <param name="item5">optional object item5</param>
        /// <param name="field6">optional object field6</param>
        /// <param name="item6">optional object item6</param>
        /// <param name="field7">optional object field7</param>
        /// <param name="item7">optional object item7</param>
        /// <param name="field8">optional object field8</param>
        /// <param name="item8">optional object item8</param>
        /// <param name="field9">optional object field9</param>
        /// <param name="item9">optional object item9</param>
        /// <param name="field10">optional object field10</param>
        /// <param name="item10">optional object item10</param>
        /// <param name="field11">optional object field11</param>
        /// <param name="item11">optional object item11</param>
        /// <param name="field12">optional object field12</param>
        /// <param name="item12">optional object item12</param>
        /// <param name="field13">optional object field13</param>
        [CustomMethod]
        [SupportByVersion("Excel", 10, 11, 12, 14, 15, 16)]
        public virtual NetOffice.ExcelApi.Range GetPivotData(object dataField, object field1, object item1, object field2, object item2, object field3, object item3, object field4, object item4, object field5, object item5, object field6, object item6, object field7, object item7, object field8, object item8, object field9, object item9, object field10, object item10, object field11, object item11, object field12, object item12, object field13)
        {
            return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.ExcelApi.Range>(this, "GetPivotData", typeof(NetOffice.ExcelApi.Range), new object[] { dataField, field1, item1, field2, item2, field3, item3, field4, item4, field5, item5, field6, item6, field7, item7, field8, item8, field9, item9, field10, item10, field11, item11, field12, item12, field13 });
        }

        /// <summary>
        /// SupportByVersion Excel 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="dataField">optional object dataField</param>
        /// <param name="field1">optional object field1</param>
        /// <param name="item1">optional object item1</param>
        /// <param name="field2">optional object field2</param>
        /// <param name="item2">optional object item2</param>
        /// <param name="field3">optional object field3</param>
        /// <param name="item3">optional object item3</param>
        /// <param name="field4">optional object field4</param>
        /// <param name="item4">optional object item4</param>
        /// <param name="field5">optional object field5</param>
        /// <param name="item5">optional object item5</param>
        /// <param name="field6">optional object field6</param>
        /// <param name="item6">optional object item6</param>
        /// <param name="field7">optional object field7</param>
        /// <param name="item7">optional object item7</param>
        /// <param name="field8">optional object field8</param>
        /// <param name="item8">optional object item8</param>
        /// <param name="field9">optional object field9</param>
        /// <param name="item9">optional object item9</param>
        /// <param name="field10">optional object field10</param>
        /// <param name="item10">optional object item10</param>
        /// <param name="field11">optional object field11</param>
        /// <param name="item11">optional object item11</param>
        /// <param name="field12">optional object field12</param>
        /// <param name="item12">optional object item12</param>
        /// <param name="field13">optional object field13</param>
        /// <param name="item13">optional object item13</param>
        [CustomMethod]
        [SupportByVersion("Excel", 10, 11, 12, 14, 15, 16)]
        public virtual NetOffice.ExcelApi.Range GetPivotData(object dataField, object field1, object item1, object field2, object item2, object field3, object item3, object field4, object item4, object field5, object item5, object field6, object item6, object field7, object item7, object field8, object item8, object field9, object item9, object field10, object item10, object field11, object item11, object field12, object item12, object field13, object item13)
        {
            return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.ExcelApi.Range>(this, "GetPivotData", typeof(NetOffice.ExcelApi.Range), new object[] { dataField, field1, item1, field2, item2, field3, item3, field4, item4, field5, item5, field6, item6, field7, item7, field8, item8, field9, item9, field10, item10, field11, item11, field12, item12, field13, item13 });
        }

        /// <summary>
        /// SupportByVersion Excel 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="dataField">optional object dataField</param>
        /// <param name="field1">optional object field1</param>
        /// <param name="item1">optional object item1</param>
        /// <param name="field2">optional object field2</param>
        /// <param name="item2">optional object item2</param>
        /// <param name="field3">optional object field3</param>
        /// <param name="item3">optional object item3</param>
        /// <param name="field4">optional object field4</param>
        /// <param name="item4">optional object item4</param>
        /// <param name="field5">optional object field5</param>
        /// <param name="item5">optional object item5</param>
        /// <param name="field6">optional object field6</param>
        /// <param name="item6">optional object item6</param>
        /// <param name="field7">optional object field7</param>
        /// <param name="item7">optional object item7</param>
        /// <param name="field8">optional object field8</param>
        /// <param name="item8">optional object item8</param>
        /// <param name="field9">optional object field9</param>
        /// <param name="item9">optional object item9</param>
        /// <param name="field10">optional object field10</param>
        /// <param name="item10">optional object item10</param>
        /// <param name="field11">optional object field11</param>
        /// <param name="item11">optional object item11</param>
        /// <param name="field12">optional object field12</param>
        /// <param name="item12">optional object item12</param>
        /// <param name="field13">optional object field13</param>
        /// <param name="item13">optional object item13</param>
        /// <param name="field14">optional object field14</param>
        [CustomMethod]
        [SupportByVersion("Excel", 10, 11, 12, 14, 15, 16)]
        public virtual NetOffice.ExcelApi.Range GetPivotData(object dataField, object field1, object item1, object field2, object item2, object field3, object item3, object field4, object item4, object field5, object item5, object field6, object item6, object field7, object item7, object field8, object item8, object field9, object item9, object field10, object item10, object field11, object item11, object field12, object item12, object field13, object item13, object field14)
        {
            return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.ExcelApi.Range>(this, "GetPivotData", typeof(NetOffice.ExcelApi.Range), new object[] { dataField, field1, item1, field2, item2, field3, item3, field4, item4, field5, item5, field6, item6, field7, item7, field8, item8, field9, item9, field10, item10, field11, item11, field12, item12, field13, item13, field14 });
        }

        /// <summary>
        /// SupportByVersion Excel 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="field">object field</param>
        /// <param name="caption">optional object caption</param>
        /// <param name="function">optional object function</param>
        [SupportByVersion("Excel", 10, 11, 12, 14, 15, 16)]
        public virtual NetOffice.ExcelApi.PivotField AddDataField(object field, object caption, object function)
        {
            return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.ExcelApi.PivotField>(this, "AddDataField", typeof(NetOffice.ExcelApi.PivotField), field, caption, function);
        }

        /// <summary>
        /// SupportByVersion Excel 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="field">object field</param>
        [CustomMethod]
        [SupportByVersion("Excel", 10, 11, 12, 14, 15, 16)]
        public virtual NetOffice.ExcelApi.PivotField AddDataField(object field)
        {
            return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.ExcelApi.PivotField>(this, "AddDataField", typeof(NetOffice.ExcelApi.PivotField), field);
        }

        /// <summary>
        /// SupportByVersion Excel 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="field">object field</param>
        /// <param name="caption">optional object caption</param>
        [CustomMethod]
        [SupportByVersion("Excel", 10, 11, 12, 14, 15, 16)]
        public virtual NetOffice.ExcelApi.PivotField AddDataField(object field, object caption)
        {
            return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.ExcelApi.PivotField>(this, "AddDataField", typeof(NetOffice.ExcelApi.PivotField), field, caption);
        }

        /// <summary>
        /// SupportByVersion Excel 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="arg1">object arg1</param>
        /// <param name="arg2">optional object arg2</param>
        /// <param name="arg3">optional object arg3</param>
        /// <param name="arg4">optional object arg4</param>
        /// <param name="arg5">optional object arg5</param>
        /// <param name="arg6">optional object arg6</param>
        /// <param name="arg7">optional object arg7</param>
        /// <param name="arg8">optional object arg8</param>
        /// <param name="arg9">optional object arg9</param>
        /// <param name="arg10">optional object arg10</param>
        /// <param name="arg11">optional object arg11</param>
        /// <param name="arg12">optional object arg12</param>
        /// <param name="arg13">optional object arg13</param>
        /// <param name="arg14">optional object arg14</param>
        /// <param name="arg15">optional object arg15</param>
        /// <param name="arg16">optional object arg16</param>
        /// <param name="arg17">optional object arg17</param>
        /// <param name="arg18">optional object arg18</param>
        /// <param name="arg19">optional object arg19</param>
        /// <param name="arg20">optional object arg20</param>
        /// <param name="arg21">optional object arg21</param>
        /// <param name="arg22">optional object arg22</param>
        /// <param name="arg23">optional object arg23</param>
        /// <param name="arg24">optional object arg24</param>
        /// <param name="arg25">optional object arg25</param>
        /// <param name="arg26">optional object arg26</param>
        /// <param name="arg27">optional object arg27</param>
        /// <param name="arg28">optional object arg28</param>
        /// <param name="arg29">optional object arg29</param>
        /// <param name="arg30">optional object arg30</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [SupportByVersion("Excel", 10, 11, 12, 14, 15, 16)]
        public virtual object Dummy15(object arg1, object arg2, object arg3, object arg4, object arg5, object arg6, object arg7, object arg8, object arg9, object arg10, object arg11, object arg12, object arg13, object arg14, object arg15, object arg16, object arg17, object arg18, object arg19, object arg20, object arg21, object arg22, object arg23, object arg24, object arg25, object arg26, object arg27, object arg28, object arg29, object arg30)
        {
            return InvokerService.InvokeInternal.ExecuteVariantMethodGet(this, "Dummy15", new object[] { arg1, arg2, arg3, arg4, arg5, arg6, arg7, arg8, arg9, arg10, arg11, arg12, arg13, arg14, arg15, arg16, arg17, arg18, arg19, arg20, arg21, arg22, arg23, arg24, arg25, arg26, arg27, arg28, arg29, arg30 });
        }

        /// <summary>
        /// SupportByVersion Excel 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="arg1">object arg1</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [CustomMethod]
        [SupportByVersion("Excel", 10, 11, 12, 14, 15, 16)]
        public virtual object Dummy15(object arg1)
        {
            return InvokerService.InvokeInternal.ExecuteVariantMethodGet(this, "Dummy15", arg1);
        }

        /// <summary>
        /// SupportByVersion Excel 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="arg1">object arg1</param>
        /// <param name="arg2">optional object arg2</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [CustomMethod]
        [SupportByVersion("Excel", 10, 11, 12, 14, 15, 16)]
        public virtual object Dummy15(object arg1, object arg2)
        {
            return InvokerService.InvokeInternal.ExecuteVariantMethodGet(this, "Dummy15", arg1, arg2);
        }

        /// <summary>
        /// SupportByVersion Excel 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="arg1">object arg1</param>
        /// <param name="arg2">optional object arg2</param>
        /// <param name="arg3">optional object arg3</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [CustomMethod]
        [SupportByVersion("Excel", 10, 11, 12, 14, 15, 16)]
        public virtual object Dummy15(object arg1, object arg2, object arg3)
        {
            return InvokerService.InvokeInternal.ExecuteVariantMethodGet(this, "Dummy15", arg1, arg2, arg3);
        }

        /// <summary>
        /// SupportByVersion Excel 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="arg1">object arg1</param>
        /// <param name="arg2">optional object arg2</param>
        /// <param name="arg3">optional object arg3</param>
        /// <param name="arg4">optional object arg4</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [CustomMethod]
        [SupportByVersion("Excel", 10, 11, 12, 14, 15, 16)]
        public virtual object Dummy15(object arg1, object arg2, object arg3, object arg4)
        {
            return InvokerService.InvokeInternal.ExecuteVariantMethodGet(this, "Dummy15", arg1, arg2, arg3, arg4);
        }

        /// <summary>
        /// SupportByVersion Excel 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="arg1">object arg1</param>
        /// <param name="arg2">optional object arg2</param>
        /// <param name="arg3">optional object arg3</param>
        /// <param name="arg4">optional object arg4</param>
        /// <param name="arg5">optional object arg5</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [CustomMethod]
        [SupportByVersion("Excel", 10, 11, 12, 14, 15, 16)]
        public virtual object Dummy15(object arg1, object arg2, object arg3, object arg4, object arg5)
        {
            return InvokerService.InvokeInternal.ExecuteVariantMethodGet(this, "Dummy15", new object[] { arg1, arg2, arg3, arg4, arg5 });
        }

        /// <summary>
        /// SupportByVersion Excel 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="arg1">object arg1</param>
        /// <param name="arg2">optional object arg2</param>
        /// <param name="arg3">optional object arg3</param>
        /// <param name="arg4">optional object arg4</param>
        /// <param name="arg5">optional object arg5</param>
        /// <param name="arg6">optional object arg6</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [CustomMethod]
        [SupportByVersion("Excel", 10, 11, 12, 14, 15, 16)]
        public virtual object Dummy15(object arg1, object arg2, object arg3, object arg4, object arg5, object arg6)
        {
            return InvokerService.InvokeInternal.ExecuteVariantMethodGet(this, "Dummy15", new object[] { arg1, arg2, arg3, arg4, arg5, arg6 });
        }

        /// <summary>
        /// SupportByVersion Excel 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="arg1">object arg1</param>
        /// <param name="arg2">optional object arg2</param>
        /// <param name="arg3">optional object arg3</param>
        /// <param name="arg4">optional object arg4</param>
        /// <param name="arg5">optional object arg5</param>
        /// <param name="arg6">optional object arg6</param>
        /// <param name="arg7">optional object arg7</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [CustomMethod]
        [SupportByVersion("Excel", 10, 11, 12, 14, 15, 16)]
        public virtual object Dummy15(object arg1, object arg2, object arg3, object arg4, object arg5, object arg6, object arg7)
        {
            return InvokerService.InvokeInternal.ExecuteVariantMethodGet(this, "Dummy15", new object[] { arg1, arg2, arg3, arg4, arg5, arg6, arg7 });
        }

        /// <summary>
        /// SupportByVersion Excel 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="arg1">object arg1</param>
        /// <param name="arg2">optional object arg2</param>
        /// <param name="arg3">optional object arg3</param>
        /// <param name="arg4">optional object arg4</param>
        /// <param name="arg5">optional object arg5</param>
        /// <param name="arg6">optional object arg6</param>
        /// <param name="arg7">optional object arg7</param>
        /// <param name="arg8">optional object arg8</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [CustomMethod]
        [SupportByVersion("Excel", 10, 11, 12, 14, 15, 16)]
        public virtual object Dummy15(object arg1, object arg2, object arg3, object arg4, object arg5, object arg6, object arg7, object arg8)
        {
            return InvokerService.InvokeInternal.ExecuteVariantMethodGet(this, "Dummy15", new object[] { arg1, arg2, arg3, arg4, arg5, arg6, arg7, arg8 });
        }

        /// <summary>
        /// SupportByVersion Excel 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="arg1">object arg1</param>
        /// <param name="arg2">optional object arg2</param>
        /// <param name="arg3">optional object arg3</param>
        /// <param name="arg4">optional object arg4</param>
        /// <param name="arg5">optional object arg5</param>
        /// <param name="arg6">optional object arg6</param>
        /// <param name="arg7">optional object arg7</param>
        /// <param name="arg8">optional object arg8</param>
        /// <param name="arg9">optional object arg9</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [CustomMethod]
        [SupportByVersion("Excel", 10, 11, 12, 14, 15, 16)]
        public virtual object Dummy15(object arg1, object arg2, object arg3, object arg4, object arg5, object arg6, object arg7, object arg8, object arg9)
        {
            return InvokerService.InvokeInternal.ExecuteVariantMethodGet(this, "Dummy15", new object[] { arg1, arg2, arg3, arg4, arg5, arg6, arg7, arg8, arg9 });
        }

        /// <summary>
        /// SupportByVersion Excel 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="arg1">object arg1</param>
        /// <param name="arg2">optional object arg2</param>
        /// <param name="arg3">optional object arg3</param>
        /// <param name="arg4">optional object arg4</param>
        /// <param name="arg5">optional object arg5</param>
        /// <param name="arg6">optional object arg6</param>
        /// <param name="arg7">optional object arg7</param>
        /// <param name="arg8">optional object arg8</param>
        /// <param name="arg9">optional object arg9</param>
        /// <param name="arg10">optional object arg10</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [CustomMethod]
        [SupportByVersion("Excel", 10, 11, 12, 14, 15, 16)]
        public virtual object Dummy15(object arg1, object arg2, object arg3, object arg4, object arg5, object arg6, object arg7, object arg8, object arg9, object arg10)
        {
            return InvokerService.InvokeInternal.ExecuteVariantMethodGet(this, "Dummy15", new object[] { arg1, arg2, arg3, arg4, arg5, arg6, arg7, arg8, arg9, arg10 });
        }

        /// <summary>
        /// SupportByVersion Excel 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="arg1">object arg1</param>
        /// <param name="arg2">optional object arg2</param>
        /// <param name="arg3">optional object arg3</param>
        /// <param name="arg4">optional object arg4</param>
        /// <param name="arg5">optional object arg5</param>
        /// <param name="arg6">optional object arg6</param>
        /// <param name="arg7">optional object arg7</param>
        /// <param name="arg8">optional object arg8</param>
        /// <param name="arg9">optional object arg9</param>
        /// <param name="arg10">optional object arg10</param>
        /// <param name="arg11">optional object arg11</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [CustomMethod]
        [SupportByVersion("Excel", 10, 11, 12, 14, 15, 16)]
        public virtual object Dummy15(object arg1, object arg2, object arg3, object arg4, object arg5, object arg6, object arg7, object arg8, object arg9, object arg10, object arg11)
        {
            return InvokerService.InvokeInternal.ExecuteVariantMethodGet(this, "Dummy15", new object[] { arg1, arg2, arg3, arg4, arg5, arg6, arg7, arg8, arg9, arg10, arg11 });
        }

        /// <summary>
        /// SupportByVersion Excel 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="arg1">object arg1</param>
        /// <param name="arg2">optional object arg2</param>
        /// <param name="arg3">optional object arg3</param>
        /// <param name="arg4">optional object arg4</param>
        /// <param name="arg5">optional object arg5</param>
        /// <param name="arg6">optional object arg6</param>
        /// <param name="arg7">optional object arg7</param>
        /// <param name="arg8">optional object arg8</param>
        /// <param name="arg9">optional object arg9</param>
        /// <param name="arg10">optional object arg10</param>
        /// <param name="arg11">optional object arg11</param>
        /// <param name="arg12">optional object arg12</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [CustomMethod]
        [SupportByVersion("Excel", 10, 11, 12, 14, 15, 16)]
        public virtual object Dummy15(object arg1, object arg2, object arg3, object arg4, object arg5, object arg6, object arg7, object arg8, object arg9, object arg10, object arg11, object arg12)
        {
            return InvokerService.InvokeInternal.ExecuteVariantMethodGet(this, "Dummy15", new object[] { arg1, arg2, arg3, arg4, arg5, arg6, arg7, arg8, arg9, arg10, arg11, arg12 });
        }

        /// <summary>
        /// SupportByVersion Excel 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="arg1">object arg1</param>
        /// <param name="arg2">optional object arg2</param>
        /// <param name="arg3">optional object arg3</param>
        /// <param name="arg4">optional object arg4</param>
        /// <param name="arg5">optional object arg5</param>
        /// <param name="arg6">optional object arg6</param>
        /// <param name="arg7">optional object arg7</param>
        /// <param name="arg8">optional object arg8</param>
        /// <param name="arg9">optional object arg9</param>
        /// <param name="arg10">optional object arg10</param>
        /// <param name="arg11">optional object arg11</param>
        /// <param name="arg12">optional object arg12</param>
        /// <param name="arg13">optional object arg13</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [CustomMethod]
        [SupportByVersion("Excel", 10, 11, 12, 14, 15, 16)]
        public virtual object Dummy15(object arg1, object arg2, object arg3, object arg4, object arg5, object arg6, object arg7, object arg8, object arg9, object arg10, object arg11, object arg12, object arg13)
        {
            return InvokerService.InvokeInternal.ExecuteVariantMethodGet(this, "Dummy15", new object[] { arg1, arg2, arg3, arg4, arg5, arg6, arg7, arg8, arg9, arg10, arg11, arg12, arg13 });
        }

        /// <summary>
        /// SupportByVersion Excel 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="arg1">object arg1</param>
        /// <param name="arg2">optional object arg2</param>
        /// <param name="arg3">optional object arg3</param>
        /// <param name="arg4">optional object arg4</param>
        /// <param name="arg5">optional object arg5</param>
        /// <param name="arg6">optional object arg6</param>
        /// <param name="arg7">optional object arg7</param>
        /// <param name="arg8">optional object arg8</param>
        /// <param name="arg9">optional object arg9</param>
        /// <param name="arg10">optional object arg10</param>
        /// <param name="arg11">optional object arg11</param>
        /// <param name="arg12">optional object arg12</param>
        /// <param name="arg13">optional object arg13</param>
        /// <param name="arg14">optional object arg14</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [CustomMethod]
        [SupportByVersion("Excel", 10, 11, 12, 14, 15, 16)]
        public virtual object Dummy15(object arg1, object arg2, object arg3, object arg4, object arg5, object arg6, object arg7, object arg8, object arg9, object arg10, object arg11, object arg12, object arg13, object arg14)
        {
            return InvokerService.InvokeInternal.ExecuteVariantMethodGet(this, "Dummy15", new object[] { arg1, arg2, arg3, arg4, arg5, arg6, arg7, arg8, arg9, arg10, arg11, arg12, arg13, arg14 });
        }

        /// <summary>
        /// SupportByVersion Excel 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="arg1">object arg1</param>
        /// <param name="arg2">optional object arg2</param>
        /// <param name="arg3">optional object arg3</param>
        /// <param name="arg4">optional object arg4</param>
        /// <param name="arg5">optional object arg5</param>
        /// <param name="arg6">optional object arg6</param>
        /// <param name="arg7">optional object arg7</param>
        /// <param name="arg8">optional object arg8</param>
        /// <param name="arg9">optional object arg9</param>
        /// <param name="arg10">optional object arg10</param>
        /// <param name="arg11">optional object arg11</param>
        /// <param name="arg12">optional object arg12</param>
        /// <param name="arg13">optional object arg13</param>
        /// <param name="arg14">optional object arg14</param>
        /// <param name="arg15">optional object arg15</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [CustomMethod]
        [SupportByVersion("Excel", 10, 11, 12, 14, 15, 16)]
        public virtual object Dummy15(object arg1, object arg2, object arg3, object arg4, object arg5, object arg6, object arg7, object arg8, object arg9, object arg10, object arg11, object arg12, object arg13, object arg14, object arg15)
        {
            return InvokerService.InvokeInternal.ExecuteVariantMethodGet(this, "Dummy15", new object[] { arg1, arg2, arg3, arg4, arg5, arg6, arg7, arg8, arg9, arg10, arg11, arg12, arg13, arg14, arg15 });
        }

        /// <summary>
        /// SupportByVersion Excel 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="arg1">object arg1</param>
        /// <param name="arg2">optional object arg2</param>
        /// <param name="arg3">optional object arg3</param>
        /// <param name="arg4">optional object arg4</param>
        /// <param name="arg5">optional object arg5</param>
        /// <param name="arg6">optional object arg6</param>
        /// <param name="arg7">optional object arg7</param>
        /// <param name="arg8">optional object arg8</param>
        /// <param name="arg9">optional object arg9</param>
        /// <param name="arg10">optional object arg10</param>
        /// <param name="arg11">optional object arg11</param>
        /// <param name="arg12">optional object arg12</param>
        /// <param name="arg13">optional object arg13</param>
        /// <param name="arg14">optional object arg14</param>
        /// <param name="arg15">optional object arg15</param>
        /// <param name="arg16">optional object arg16</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [CustomMethod]
        [SupportByVersion("Excel", 10, 11, 12, 14, 15, 16)]
        public virtual object Dummy15(object arg1, object arg2, object arg3, object arg4, object arg5, object arg6, object arg7, object arg8, object arg9, object arg10, object arg11, object arg12, object arg13, object arg14, object arg15, object arg16)
        {
            return InvokerService.InvokeInternal.ExecuteVariantMethodGet(this, "Dummy15", new object[] { arg1, arg2, arg3, arg4, arg5, arg6, arg7, arg8, arg9, arg10, arg11, arg12, arg13, arg14, arg15, arg16 });
        }

        /// <summary>
        /// SupportByVersion Excel 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="arg1">object arg1</param>
        /// <param name="arg2">optional object arg2</param>
        /// <param name="arg3">optional object arg3</param>
        /// <param name="arg4">optional object arg4</param>
        /// <param name="arg5">optional object arg5</param>
        /// <param name="arg6">optional object arg6</param>
        /// <param name="arg7">optional object arg7</param>
        /// <param name="arg8">optional object arg8</param>
        /// <param name="arg9">optional object arg9</param>
        /// <param name="arg10">optional object arg10</param>
        /// <param name="arg11">optional object arg11</param>
        /// <param name="arg12">optional object arg12</param>
        /// <param name="arg13">optional object arg13</param>
        /// <param name="arg14">optional object arg14</param>
        /// <param name="arg15">optional object arg15</param>
        /// <param name="arg16">optional object arg16</param>
        /// <param name="arg17">optional object arg17</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [CustomMethod]
        [SupportByVersion("Excel", 10, 11, 12, 14, 15, 16)]
        public virtual object Dummy15(object arg1, object arg2, object arg3, object arg4, object arg5, object arg6, object arg7, object arg8, object arg9, object arg10, object arg11, object arg12, object arg13, object arg14, object arg15, object arg16, object arg17)
        {
            return InvokerService.InvokeInternal.ExecuteVariantMethodGet(this, "Dummy15", new object[] { arg1, arg2, arg3, arg4, arg5, arg6, arg7, arg8, arg9, arg10, arg11, arg12, arg13, arg14, arg15, arg16, arg17 });
        }

        /// <summary>
        /// SupportByVersion Excel 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="arg1">object arg1</param>
        /// <param name="arg2">optional object arg2</param>
        /// <param name="arg3">optional object arg3</param>
        /// <param name="arg4">optional object arg4</param>
        /// <param name="arg5">optional object arg5</param>
        /// <param name="arg6">optional object arg6</param>
        /// <param name="arg7">optional object arg7</param>
        /// <param name="arg8">optional object arg8</param>
        /// <param name="arg9">optional object arg9</param>
        /// <param name="arg10">optional object arg10</param>
        /// <param name="arg11">optional object arg11</param>
        /// <param name="arg12">optional object arg12</param>
        /// <param name="arg13">optional object arg13</param>
        /// <param name="arg14">optional object arg14</param>
        /// <param name="arg15">optional object arg15</param>
        /// <param name="arg16">optional object arg16</param>
        /// <param name="arg17">optional object arg17</param>
        /// <param name="arg18">optional object arg18</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [CustomMethod]
        [SupportByVersion("Excel", 10, 11, 12, 14, 15, 16)]
        public virtual object Dummy15(object arg1, object arg2, object arg3, object arg4, object arg5, object arg6, object arg7, object arg8, object arg9, object arg10, object arg11, object arg12, object arg13, object arg14, object arg15, object arg16, object arg17, object arg18)
        {
            return InvokerService.InvokeInternal.ExecuteVariantMethodGet(this, "Dummy15", new object[] { arg1, arg2, arg3, arg4, arg5, arg6, arg7, arg8, arg9, arg10, arg11, arg12, arg13, arg14, arg15, arg16, arg17, arg18 });
        }

        /// <summary>
        /// SupportByVersion Excel 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="arg1">object arg1</param>
        /// <param name="arg2">optional object arg2</param>
        /// <param name="arg3">optional object arg3</param>
        /// <param name="arg4">optional object arg4</param>
        /// <param name="arg5">optional object arg5</param>
        /// <param name="arg6">optional object arg6</param>
        /// <param name="arg7">optional object arg7</param>
        /// <param name="arg8">optional object arg8</param>
        /// <param name="arg9">optional object arg9</param>
        /// <param name="arg10">optional object arg10</param>
        /// <param name="arg11">optional object arg11</param>
        /// <param name="arg12">optional object arg12</param>
        /// <param name="arg13">optional object arg13</param>
        /// <param name="arg14">optional object arg14</param>
        /// <param name="arg15">optional object arg15</param>
        /// <param name="arg16">optional object arg16</param>
        /// <param name="arg17">optional object arg17</param>
        /// <param name="arg18">optional object arg18</param>
        /// <param name="arg19">optional object arg19</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [CustomMethod]
        [SupportByVersion("Excel", 10, 11, 12, 14, 15, 16)]
        public virtual object Dummy15(object arg1, object arg2, object arg3, object arg4, object arg5, object arg6, object arg7, object arg8, object arg9, object arg10, object arg11, object arg12, object arg13, object arg14, object arg15, object arg16, object arg17, object arg18, object arg19)
        {
            return InvokerService.InvokeInternal.ExecuteVariantMethodGet(this, "Dummy15", new object[] { arg1, arg2, arg3, arg4, arg5, arg6, arg7, arg8, arg9, arg10, arg11, arg12, arg13, arg14, arg15, arg16, arg17, arg18, arg19 });
        }

        /// <summary>
        /// SupportByVersion Excel 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="arg1">object arg1</param>
        /// <param name="arg2">optional object arg2</param>
        /// <param name="arg3">optional object arg3</param>
        /// <param name="arg4">optional object arg4</param>
        /// <param name="arg5">optional object arg5</param>
        /// <param name="arg6">optional object arg6</param>
        /// <param name="arg7">optional object arg7</param>
        /// <param name="arg8">optional object arg8</param>
        /// <param name="arg9">optional object arg9</param>
        /// <param name="arg10">optional object arg10</param>
        /// <param name="arg11">optional object arg11</param>
        /// <param name="arg12">optional object arg12</param>
        /// <param name="arg13">optional object arg13</param>
        /// <param name="arg14">optional object arg14</param>
        /// <param name="arg15">optional object arg15</param>
        /// <param name="arg16">optional object arg16</param>
        /// <param name="arg17">optional object arg17</param>
        /// <param name="arg18">optional object arg18</param>
        /// <param name="arg19">optional object arg19</param>
        /// <param name="arg20">optional object arg20</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [CustomMethod]
        [SupportByVersion("Excel", 10, 11, 12, 14, 15, 16)]
        public virtual object Dummy15(object arg1, object arg2, object arg3, object arg4, object arg5, object arg6, object arg7, object arg8, object arg9, object arg10, object arg11, object arg12, object arg13, object arg14, object arg15, object arg16, object arg17, object arg18, object arg19, object arg20)
        {
            return InvokerService.InvokeInternal.ExecuteVariantMethodGet(this, "Dummy15", new object[] { arg1, arg2, arg3, arg4, arg5, arg6, arg7, arg8, arg9, arg10, arg11, arg12, arg13, arg14, arg15, arg16, arg17, arg18, arg19, arg20 });
        }

        /// <summary>
        /// SupportByVersion Excel 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="arg1">object arg1</param>
        /// <param name="arg2">optional object arg2</param>
        /// <param name="arg3">optional object arg3</param>
        /// <param name="arg4">optional object arg4</param>
        /// <param name="arg5">optional object arg5</param>
        /// <param name="arg6">optional object arg6</param>
        /// <param name="arg7">optional object arg7</param>
        /// <param name="arg8">optional object arg8</param>
        /// <param name="arg9">optional object arg9</param>
        /// <param name="arg10">optional object arg10</param>
        /// <param name="arg11">optional object arg11</param>
        /// <param name="arg12">optional object arg12</param>
        /// <param name="arg13">optional object arg13</param>
        /// <param name="arg14">optional object arg14</param>
        /// <param name="arg15">optional object arg15</param>
        /// <param name="arg16">optional object arg16</param>
        /// <param name="arg17">optional object arg17</param>
        /// <param name="arg18">optional object arg18</param>
        /// <param name="arg19">optional object arg19</param>
        /// <param name="arg20">optional object arg20</param>
        /// <param name="arg21">optional object arg21</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [CustomMethod]
        [SupportByVersion("Excel", 10, 11, 12, 14, 15, 16)]
        public virtual object Dummy15(object arg1, object arg2, object arg3, object arg4, object arg5, object arg6, object arg7, object arg8, object arg9, object arg10, object arg11, object arg12, object arg13, object arg14, object arg15, object arg16, object arg17, object arg18, object arg19, object arg20, object arg21)
        {
            return InvokerService.InvokeInternal.ExecuteVariantMethodGet(this, "Dummy15", new object[] { arg1, arg2, arg3, arg4, arg5, arg6, arg7, arg8, arg9, arg10, arg11, arg12, arg13, arg14, arg15, arg16, arg17, arg18, arg19, arg20, arg21 });
        }

        /// <summary>
        /// SupportByVersion Excel 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="arg1">object arg1</param>
        /// <param name="arg2">optional object arg2</param>
        /// <param name="arg3">optional object arg3</param>
        /// <param name="arg4">optional object arg4</param>
        /// <param name="arg5">optional object arg5</param>
        /// <param name="arg6">optional object arg6</param>
        /// <param name="arg7">optional object arg7</param>
        /// <param name="arg8">optional object arg8</param>
        /// <param name="arg9">optional object arg9</param>
        /// <param name="arg10">optional object arg10</param>
        /// <param name="arg11">optional object arg11</param>
        /// <param name="arg12">optional object arg12</param>
        /// <param name="arg13">optional object arg13</param>
        /// <param name="arg14">optional object arg14</param>
        /// <param name="arg15">optional object arg15</param>
        /// <param name="arg16">optional object arg16</param>
        /// <param name="arg17">optional object arg17</param>
        /// <param name="arg18">optional object arg18</param>
        /// <param name="arg19">optional object arg19</param>
        /// <param name="arg20">optional object arg20</param>
        /// <param name="arg21">optional object arg21</param>
        /// <param name="arg22">optional object arg22</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [CustomMethod]
        [SupportByVersion("Excel", 10, 11, 12, 14, 15, 16)]
        public virtual object Dummy15(object arg1, object arg2, object arg3, object arg4, object arg5, object arg6, object arg7, object arg8, object arg9, object arg10, object arg11, object arg12, object arg13, object arg14, object arg15, object arg16, object arg17, object arg18, object arg19, object arg20, object arg21, object arg22)
        {
            return InvokerService.InvokeInternal.ExecuteVariantMethodGet(this, "Dummy15", new object[] { arg1, arg2, arg3, arg4, arg5, arg6, arg7, arg8, arg9, arg10, arg11, arg12, arg13, arg14, arg15, arg16, arg17, arg18, arg19, arg20, arg21, arg22 });
        }

        /// <summary>
        /// SupportByVersion Excel 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="arg1">object arg1</param>
        /// <param name="arg2">optional object arg2</param>
        /// <param name="arg3">optional object arg3</param>
        /// <param name="arg4">optional object arg4</param>
        /// <param name="arg5">optional object arg5</param>
        /// <param name="arg6">optional object arg6</param>
        /// <param name="arg7">optional object arg7</param>
        /// <param name="arg8">optional object arg8</param>
        /// <param name="arg9">optional object arg9</param>
        /// <param name="arg10">optional object arg10</param>
        /// <param name="arg11">optional object arg11</param>
        /// <param name="arg12">optional object arg12</param>
        /// <param name="arg13">optional object arg13</param>
        /// <param name="arg14">optional object arg14</param>
        /// <param name="arg15">optional object arg15</param>
        /// <param name="arg16">optional object arg16</param>
        /// <param name="arg17">optional object arg17</param>
        /// <param name="arg18">optional object arg18</param>
        /// <param name="arg19">optional object arg19</param>
        /// <param name="arg20">optional object arg20</param>
        /// <param name="arg21">optional object arg21</param>
        /// <param name="arg22">optional object arg22</param>
        /// <param name="arg23">optional object arg23</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [CustomMethod]
        [SupportByVersion("Excel", 10, 11, 12, 14, 15, 16)]
        public virtual object Dummy15(object arg1, object arg2, object arg3, object arg4, object arg5, object arg6, object arg7, object arg8, object arg9, object arg10, object arg11, object arg12, object arg13, object arg14, object arg15, object arg16, object arg17, object arg18, object arg19, object arg20, object arg21, object arg22, object arg23)
        {
            return InvokerService.InvokeInternal.ExecuteVariantMethodGet(this, "Dummy15", new object[] { arg1, arg2, arg3, arg4, arg5, arg6, arg7, arg8, arg9, arg10, arg11, arg12, arg13, arg14, arg15, arg16, arg17, arg18, arg19, arg20, arg21, arg22, arg23 });
        }

        /// <summary>
        /// SupportByVersion Excel 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="arg1">object arg1</param>
        /// <param name="arg2">optional object arg2</param>
        /// <param name="arg3">optional object arg3</param>
        /// <param name="arg4">optional object arg4</param>
        /// <param name="arg5">optional object arg5</param>
        /// <param name="arg6">optional object arg6</param>
        /// <param name="arg7">optional object arg7</param>
        /// <param name="arg8">optional object arg8</param>
        /// <param name="arg9">optional object arg9</param>
        /// <param name="arg10">optional object arg10</param>
        /// <param name="arg11">optional object arg11</param>
        /// <param name="arg12">optional object arg12</param>
        /// <param name="arg13">optional object arg13</param>
        /// <param name="arg14">optional object arg14</param>
        /// <param name="arg15">optional object arg15</param>
        /// <param name="arg16">optional object arg16</param>
        /// <param name="arg17">optional object arg17</param>
        /// <param name="arg18">optional object arg18</param>
        /// <param name="arg19">optional object arg19</param>
        /// <param name="arg20">optional object arg20</param>
        /// <param name="arg21">optional object arg21</param>
        /// <param name="arg22">optional object arg22</param>
        /// <param name="arg23">optional object arg23</param>
        /// <param name="arg24">optional object arg24</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [CustomMethod]
        [SupportByVersion("Excel", 10, 11, 12, 14, 15, 16)]
        public virtual object Dummy15(object arg1, object arg2, object arg3, object arg4, object arg5, object arg6, object arg7, object arg8, object arg9, object arg10, object arg11, object arg12, object arg13, object arg14, object arg15, object arg16, object arg17, object arg18, object arg19, object arg20, object arg21, object arg22, object arg23, object arg24)
        {
            return InvokerService.InvokeInternal.ExecuteVariantMethodGet(this, "Dummy15", new object[] { arg1, arg2, arg3, arg4, arg5, arg6, arg7, arg8, arg9, arg10, arg11, arg12, arg13, arg14, arg15, arg16, arg17, arg18, arg19, arg20, arg21, arg22, arg23, arg24 });
        }

        /// <summary>
        /// SupportByVersion Excel 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="arg1">object arg1</param>
        /// <param name="arg2">optional object arg2</param>
        /// <param name="arg3">optional object arg3</param>
        /// <param name="arg4">optional object arg4</param>
        /// <param name="arg5">optional object arg5</param>
        /// <param name="arg6">optional object arg6</param>
        /// <param name="arg7">optional object arg7</param>
        /// <param name="arg8">optional object arg8</param>
        /// <param name="arg9">optional object arg9</param>
        /// <param name="arg10">optional object arg10</param>
        /// <param name="arg11">optional object arg11</param>
        /// <param name="arg12">optional object arg12</param>
        /// <param name="arg13">optional object arg13</param>
        /// <param name="arg14">optional object arg14</param>
        /// <param name="arg15">optional object arg15</param>
        /// <param name="arg16">optional object arg16</param>
        /// <param name="arg17">optional object arg17</param>
        /// <param name="arg18">optional object arg18</param>
        /// <param name="arg19">optional object arg19</param>
        /// <param name="arg20">optional object arg20</param>
        /// <param name="arg21">optional object arg21</param>
        /// <param name="arg22">optional object arg22</param>
        /// <param name="arg23">optional object arg23</param>
        /// <param name="arg24">optional object arg24</param>
        /// <param name="arg25">optional object arg25</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [CustomMethod]
        [SupportByVersion("Excel", 10, 11, 12, 14, 15, 16)]
        public virtual object Dummy15(object arg1, object arg2, object arg3, object arg4, object arg5, object arg6, object arg7, object arg8, object arg9, object arg10, object arg11, object arg12, object arg13, object arg14, object arg15, object arg16, object arg17, object arg18, object arg19, object arg20, object arg21, object arg22, object arg23, object arg24, object arg25)
        {
            return InvokerService.InvokeInternal.ExecuteVariantMethodGet(this, "Dummy15", new object[] { arg1, arg2, arg3, arg4, arg5, arg6, arg7, arg8, arg9, arg10, arg11, arg12, arg13, arg14, arg15, arg16, arg17, arg18, arg19, arg20, arg21, arg22, arg23, arg24, arg25 });
        }

        /// <summary>
        /// SupportByVersion Excel 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="arg1">object arg1</param>
        /// <param name="arg2">optional object arg2</param>
        /// <param name="arg3">optional object arg3</param>
        /// <param name="arg4">optional object arg4</param>
        /// <param name="arg5">optional object arg5</param>
        /// <param name="arg6">optional object arg6</param>
        /// <param name="arg7">optional object arg7</param>
        /// <param name="arg8">optional object arg8</param>
        /// <param name="arg9">optional object arg9</param>
        /// <param name="arg10">optional object arg10</param>
        /// <param name="arg11">optional object arg11</param>
        /// <param name="arg12">optional object arg12</param>
        /// <param name="arg13">optional object arg13</param>
        /// <param name="arg14">optional object arg14</param>
        /// <param name="arg15">optional object arg15</param>
        /// <param name="arg16">optional object arg16</param>
        /// <param name="arg17">optional object arg17</param>
        /// <param name="arg18">optional object arg18</param>
        /// <param name="arg19">optional object arg19</param>
        /// <param name="arg20">optional object arg20</param>
        /// <param name="arg21">optional object arg21</param>
        /// <param name="arg22">optional object arg22</param>
        /// <param name="arg23">optional object arg23</param>
        /// <param name="arg24">optional object arg24</param>
        /// <param name="arg25">optional object arg25</param>
        /// <param name="arg26">optional object arg26</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [CustomMethod]
        [SupportByVersion("Excel", 10, 11, 12, 14, 15, 16)]
        public virtual object Dummy15(object arg1, object arg2, object arg3, object arg4, object arg5, object arg6, object arg7, object arg8, object arg9, object arg10, object arg11, object arg12, object arg13, object arg14, object arg15, object arg16, object arg17, object arg18, object arg19, object arg20, object arg21, object arg22, object arg23, object arg24, object arg25, object arg26)
        {
            return InvokerService.InvokeInternal.ExecuteVariantMethodGet(this, "Dummy15", new object[] { arg1, arg2, arg3, arg4, arg5, arg6, arg7, arg8, arg9, arg10, arg11, arg12, arg13, arg14, arg15, arg16, arg17, arg18, arg19, arg20, arg21, arg22, arg23, arg24, arg25, arg26 });
        }

        /// <summary>
        /// SupportByVersion Excel 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="arg1">object arg1</param>
        /// <param name="arg2">optional object arg2</param>
        /// <param name="arg3">optional object arg3</param>
        /// <param name="arg4">optional object arg4</param>
        /// <param name="arg5">optional object arg5</param>
        /// <param name="arg6">optional object arg6</param>
        /// <param name="arg7">optional object arg7</param>
        /// <param name="arg8">optional object arg8</param>
        /// <param name="arg9">optional object arg9</param>
        /// <param name="arg10">optional object arg10</param>
        /// <param name="arg11">optional object arg11</param>
        /// <param name="arg12">optional object arg12</param>
        /// <param name="arg13">optional object arg13</param>
        /// <param name="arg14">optional object arg14</param>
        /// <param name="arg15">optional object arg15</param>
        /// <param name="arg16">optional object arg16</param>
        /// <param name="arg17">optional object arg17</param>
        /// <param name="arg18">optional object arg18</param>
        /// <param name="arg19">optional object arg19</param>
        /// <param name="arg20">optional object arg20</param>
        /// <param name="arg21">optional object arg21</param>
        /// <param name="arg22">optional object arg22</param>
        /// <param name="arg23">optional object arg23</param>
        /// <param name="arg24">optional object arg24</param>
        /// <param name="arg25">optional object arg25</param>
        /// <param name="arg26">optional object arg26</param>
        /// <param name="arg27">optional object arg27</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [CustomMethod]
        [SupportByVersion("Excel", 10, 11, 12, 14, 15, 16)]
        public virtual object Dummy15(object arg1, object arg2, object arg3, object arg4, object arg5, object arg6, object arg7, object arg8, object arg9, object arg10, object arg11, object arg12, object arg13, object arg14, object arg15, object arg16, object arg17, object arg18, object arg19, object arg20, object arg21, object arg22, object arg23, object arg24, object arg25, object arg26, object arg27)
        {
            return InvokerService.InvokeInternal.ExecuteVariantMethodGet(this, "Dummy15", new object[] { arg1, arg2, arg3, arg4, arg5, arg6, arg7, arg8, arg9, arg10, arg11, arg12, arg13, arg14, arg15, arg16, arg17, arg18, arg19, arg20, arg21, arg22, arg23, arg24, arg25, arg26, arg27 });
        }

        /// <summary>
        /// SupportByVersion Excel 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="arg1">object arg1</param>
        /// <param name="arg2">optional object arg2</param>
        /// <param name="arg3">optional object arg3</param>
        /// <param name="arg4">optional object arg4</param>
        /// <param name="arg5">optional object arg5</param>
        /// <param name="arg6">optional object arg6</param>
        /// <param name="arg7">optional object arg7</param>
        /// <param name="arg8">optional object arg8</param>
        /// <param name="arg9">optional object arg9</param>
        /// <param name="arg10">optional object arg10</param>
        /// <param name="arg11">optional object arg11</param>
        /// <param name="arg12">optional object arg12</param>
        /// <param name="arg13">optional object arg13</param>
        /// <param name="arg14">optional object arg14</param>
        /// <param name="arg15">optional object arg15</param>
        /// <param name="arg16">optional object arg16</param>
        /// <param name="arg17">optional object arg17</param>
        /// <param name="arg18">optional object arg18</param>
        /// <param name="arg19">optional object arg19</param>
        /// <param name="arg20">optional object arg20</param>
        /// <param name="arg21">optional object arg21</param>
        /// <param name="arg22">optional object arg22</param>
        /// <param name="arg23">optional object arg23</param>
        /// <param name="arg24">optional object arg24</param>
        /// <param name="arg25">optional object arg25</param>
        /// <param name="arg26">optional object arg26</param>
        /// <param name="arg27">optional object arg27</param>
        /// <param name="arg28">optional object arg28</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [CustomMethod]
        [SupportByVersion("Excel", 10, 11, 12, 14, 15, 16)]
        public virtual object Dummy15(object arg1, object arg2, object arg3, object arg4, object arg5, object arg6, object arg7, object arg8, object arg9, object arg10, object arg11, object arg12, object arg13, object arg14, object arg15, object arg16, object arg17, object arg18, object arg19, object arg20, object arg21, object arg22, object arg23, object arg24, object arg25, object arg26, object arg27, object arg28)
        {
            return InvokerService.InvokeInternal.ExecuteVariantMethodGet(this, "Dummy15", new object[] { arg1, arg2, arg3, arg4, arg5, arg6, arg7, arg8, arg9, arg10, arg11, arg12, arg13, arg14, arg15, arg16, arg17, arg18, arg19, arg20, arg21, arg22, arg23, arg24, arg25, arg26, arg27, arg28 });
        }

        /// <summary>
        /// SupportByVersion Excel 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="arg1">object arg1</param>
        /// <param name="arg2">optional object arg2</param>
        /// <param name="arg3">optional object arg3</param>
        /// <param name="arg4">optional object arg4</param>
        /// <param name="arg5">optional object arg5</param>
        /// <param name="arg6">optional object arg6</param>
        /// <param name="arg7">optional object arg7</param>
        /// <param name="arg8">optional object arg8</param>
        /// <param name="arg9">optional object arg9</param>
        /// <param name="arg10">optional object arg10</param>
        /// <param name="arg11">optional object arg11</param>
        /// <param name="arg12">optional object arg12</param>
        /// <param name="arg13">optional object arg13</param>
        /// <param name="arg14">optional object arg14</param>
        /// <param name="arg15">optional object arg15</param>
        /// <param name="arg16">optional object arg16</param>
        /// <param name="arg17">optional object arg17</param>
        /// <param name="arg18">optional object arg18</param>
        /// <param name="arg19">optional object arg19</param>
        /// <param name="arg20">optional object arg20</param>
        /// <param name="arg21">optional object arg21</param>
        /// <param name="arg22">optional object arg22</param>
        /// <param name="arg23">optional object arg23</param>
        /// <param name="arg24">optional object arg24</param>
        /// <param name="arg25">optional object arg25</param>
        /// <param name="arg26">optional object arg26</param>
        /// <param name="arg27">optional object arg27</param>
        /// <param name="arg28">optional object arg28</param>
        /// <param name="arg29">optional object arg29</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [CustomMethod]
        [SupportByVersion("Excel", 10, 11, 12, 14, 15, 16)]
        public virtual object Dummy15(object arg1, object arg2, object arg3, object arg4, object arg5, object arg6, object arg7, object arg8, object arg9, object arg10, object arg11, object arg12, object arg13, object arg14, object arg15, object arg16, object arg17, object arg18, object arg19, object arg20, object arg21, object arg22, object arg23, object arg24, object arg25, object arg26, object arg27, object arg28, object arg29)
        {
            return InvokerService.InvokeInternal.ExecuteVariantMethodGet(this, "Dummy15", new object[] { arg1, arg2, arg3, arg4, arg5, arg6, arg7, arg8, arg9, arg10, arg11, arg12, arg13, arg14, arg15, arg16, arg17, arg18, arg19, arg20, arg21, arg22, arg23, arg24, arg25, arg26, arg27, arg28, arg29 });
        }

        /// <summary>
        /// SupportByVersion Excel 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="file">string file</param>
        /// <param name="measures">optional object measures</param>
        /// <param name="levels">optional object levels</param>
        /// <param name="members">optional object members</param>
        /// <param name="properties">optional object properties</param>
        [SupportByVersion("Excel", 10, 11, 12, 14, 15, 16)]
        public virtual string CreateCubeFile(string file, object measures, object levels, object members, object properties)
        {
            return InvokerService.InvokeInternal.ExecuteStringMethodGet(this, "CreateCubeFile", new object[] { file, measures, levels, members, properties });
        }

        /// <summary>
        /// SupportByVersion Excel 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="file">string file</param>
        [CustomMethod]
        [SupportByVersion("Excel", 10, 11, 12, 14, 15, 16)]
        public virtual string CreateCubeFile(string file)
        {
            return InvokerService.InvokeInternal.ExecuteStringMethodGet(this, "CreateCubeFile", file);
        }

        /// <summary>
        /// SupportByVersion Excel 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="file">string file</param>
        /// <param name="measures">optional object measures</param>
        [CustomMethod]
        [SupportByVersion("Excel", 10, 11, 12, 14, 15, 16)]
        public virtual string CreateCubeFile(string file, object measures)
        {
            return InvokerService.InvokeInternal.ExecuteStringMethodGet(this, "CreateCubeFile", file, measures);
        }

        /// <summary>
        /// SupportByVersion Excel 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="file">string file</param>
        /// <param name="measures">optional object measures</param>
        /// <param name="levels">optional object levels</param>
        [CustomMethod]
        [SupportByVersion("Excel", 10, 11, 12, 14, 15, 16)]
        public virtual string CreateCubeFile(string file, object measures, object levels)
        {
            return InvokerService.InvokeInternal.ExecuteStringMethodGet(this, "CreateCubeFile", file, measures, levels);
        }

        /// <summary>
        /// SupportByVersion Excel 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="file">string file</param>
        /// <param name="measures">optional object measures</param>
        /// <param name="levels">optional object levels</param>
        /// <param name="members">optional object members</param>
        [CustomMethod]
        [SupportByVersion("Excel", 10, 11, 12, 14, 15, 16)]
        public virtual string CreateCubeFile(string file, object measures, object levels, object members)
        {
            return InvokerService.InvokeInternal.ExecuteStringMethodGet(this, "CreateCubeFile", file, measures, levels, members);
        }

        /// <summary>
        /// SupportByVersion Excel 12, 14, 15, 16
        /// </summary>
        [SupportByVersion("Excel", 12, 14, 15, 16)]
        public virtual Int32 ClearTable()
        {
            return InvokerService.InvokeInternal.ExecuteInt32MethodGet(this, "ClearTable");
        }

        /// <summary>
        /// SupportByVersion Excel 12, 14, 15, 16
        /// </summary>
        /// <param name="rowLayout">NetOffice.ExcelApi.Enums.XlLayoutRowType rowLayout</param>
        [SupportByVersion("Excel", 12, 14, 15, 16)]
        public virtual Int32 RowAxisLayout(NetOffice.ExcelApi.Enums.XlLayoutRowType rowLayout)
        {
            return InvokerService.InvokeInternal.ExecuteInt32MethodGet(this, "RowAxisLayout", rowLayout);
        }

        /// <summary>
        /// SupportByVersion Excel 12, 14, 15, 16
        /// </summary>
        /// <param name="location">NetOffice.ExcelApi.Enums.xLSubtototalLocationType location</param>
        [SupportByVersion("Excel", 12, 14, 15, 16)]
        public virtual Int32 SubtotalLocation(NetOffice.ExcelApi.Enums.xLSubtototalLocationType location)
        {
            return InvokerService.InvokeInternal.ExecuteInt32MethodGet(this, "SubtotalLocation", location);
        }

        /// <summary>
        /// SupportByVersion Excel 12, 14, 15, 16
        /// </summary>
        [SupportByVersion("Excel", 12, 14, 15, 16)]
        public virtual Int32 ClearAllFilters()
        {
            return InvokerService.InvokeInternal.ExecuteInt32MethodGet(this, "ClearAllFilters");
        }

        /// <summary>
        /// SupportByVersion Excel 12, 14, 15, 16
        /// </summary>
        /// <param name="convertFilters">bool convertFilters</param>
        [SupportByVersion("Excel", 12, 14, 15, 16)]
        public virtual Int32 ConvertToFormulas(bool convertFilters)
        {
            return InvokerService.InvokeInternal.ExecuteInt32MethodGet(this, "ConvertToFormulas", convertFilters);
        }

        /// <summary>
        /// SupportByVersion Excel 12, 14, 15, 16
        /// </summary>
        /// <param name="conn">NetOffice.ExcelApi.WorkbookConnection conn</param>
        [SupportByVersion("Excel", 12, 14, 15, 16)]
        public virtual Int32 ChangeConnection(NetOffice.ExcelApi.WorkbookConnection conn)
        {
            return InvokerService.InvokeInternal.ExecuteInt32MethodGet(this, "ChangeConnection", conn);
        }

        /// <summary>
        /// SupportByVersion Excel 12, 14, 15, 16
        /// </summary>
        /// <param name="pivotCache">object pivotCache</param>
        [SupportByVersion("Excel", 12, 14, 15, 16)]
        public virtual Int32 ChangePivotCache(object pivotCache)
        {
            return InvokerService.InvokeInternal.ExecuteInt32MethodGet(this, "ChangePivotCache", pivotCache);
        }

        /// <summary>
        /// SupportByVersion Excel 14, 15, 16
        /// </summary>
        [SupportByVersion("Excel", 14, 15, 16)]
        public virtual Int32 AllocateChanges()
        {
            return InvokerService.InvokeInternal.ExecuteInt32MethodGet(this, "AllocateChanges");
        }

        /// <summary>
        /// SupportByVersion Excel 14, 15, 16
        /// </summary>
        [SupportByVersion("Excel", 14, 15, 16)]
        public virtual Int32 CommitChanges()
        {
            return InvokerService.InvokeInternal.ExecuteInt32MethodGet(this, "CommitChanges");
        }

        /// <summary>
        /// SupportByVersion Excel 14, 15, 16
        /// </summary>
        [SupportByVersion("Excel", 14, 15, 16)]
        public virtual Int32 DiscardChanges()
        {
            return InvokerService.InvokeInternal.ExecuteInt32MethodGet(this, "DiscardChanges");
        }

        /// <summary>
        /// SupportByVersion Excel 14, 15, 16
        /// </summary>
        [SupportByVersion("Excel", 14, 15, 16)]
        public virtual Int32 RefreshDataSourceValues()
        {
            return InvokerService.InvokeInternal.ExecuteInt32MethodGet(this, "RefreshDataSourceValues");
        }

        /// <summary>
        /// SupportByVersion Excel 14, 15, 16
        /// </summary>
        /// <param name="repeat">NetOffice.ExcelApi.Enums.XlPivotFieldRepeatLabels repeat</param>
        [SupportByVersion("Excel", 14, 15, 16)]
        public virtual Int32 RepeatAllLabels(NetOffice.ExcelApi.Enums.XlPivotFieldRepeatLabels repeat)
        {
            return InvokerService.InvokeInternal.ExecuteInt32MethodGet(this, "RepeatAllLabels", repeat);
        }

        /// <summary>
        /// SupportByVersion Excel 15,16
        /// </summary>
        /// <param name="rowline">optional object rowline</param>
        /// <param name="columnline">optional object columnline</param>
        [SupportByVersion("Excel", 15, 16)]
        public virtual NetOffice.ExcelApi.PivotValueCell PivotValueCell(object rowline, object columnline)
        {
            return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.ExcelApi.PivotValueCell>(this, "PivotValueCell", typeof(NetOffice.ExcelApi.PivotValueCell), rowline, columnline);
        }

        /// <summary>
        /// SupportByVersion Excel 15,16
        /// </summary>
        [CustomMethod]
        [SupportByVersion("Excel", 15, 16)]
        public virtual NetOffice.ExcelApi.PivotValueCell PivotValueCell()
        {
            return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.ExcelApi.PivotValueCell>(this, "PivotValueCell", typeof(NetOffice.ExcelApi.PivotValueCell));
        }

        /// <summary>
        /// SupportByVersion Excel 15,16
        /// </summary>
        /// <param name="rowline">optional object rowline</param>
        [CustomMethod]
        [SupportByVersion("Excel", 15, 16)]
        public virtual NetOffice.ExcelApi.PivotValueCell PivotValueCell(object rowline)
        {
            return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.ExcelApi.PivotValueCell>(this, "PivotValueCell", typeof(NetOffice.ExcelApi.PivotValueCell), rowline);
        }

        /// <summary>
        /// SupportByVersion Excel 15,16
        /// </summary>
        /// <param name="pivotItem">NetOffice.ExcelApi.PivotItem pivotItem</param>
        /// <param name="pivotLine">optional object pivotLine</param>
        [SupportByVersion("Excel", 15, 16)]
        public virtual Int32 DrillDown(NetOffice.ExcelApi.PivotItem pivotItem, object pivotLine)
        {
            return InvokerService.InvokeInternal.ExecuteInt32MethodGet(this, "DrillDown", pivotItem, pivotLine);
        }

        /// <summary>
        /// SupportByVersion Excel 15,16
        /// </summary>
        /// <param name="pivotItem">NetOffice.ExcelApi.PivotItem pivotItem</param>
        [CustomMethod]
        [SupportByVersion("Excel", 15, 16)]
        public virtual Int32 DrillDown(NetOffice.ExcelApi.PivotItem pivotItem)
        {
            return InvokerService.InvokeInternal.ExecuteInt32MethodGet(this, "DrillDown", pivotItem);
        }

        /// <summary>
        /// SupportByVersion Excel 15,16
        /// </summary>
        /// <param name="pivotItem">NetOffice.ExcelApi.PivotItem pivotItem</param>
        /// <param name="pivotLine">optional object pivotLine</param>
        /// <param name="levelUniqueName">optional object levelUniqueName</param>
        [SupportByVersion("Excel", 15, 16)]
        public virtual Int32 DrillUp(NetOffice.ExcelApi.PivotItem pivotItem, object pivotLine, object levelUniqueName)
        {
            return InvokerService.InvokeInternal.ExecuteInt32MethodGet(this, "DrillUp", pivotItem, pivotLine, levelUniqueName);
        }

        /// <summary>
        /// SupportByVersion Excel 15,16
        /// </summary>
        /// <param name="pivotItem">NetOffice.ExcelApi.PivotItem pivotItem</param>
        [CustomMethod]
        [SupportByVersion("Excel", 15, 16)]
        public virtual Int32 DrillUp(NetOffice.ExcelApi.PivotItem pivotItem)
        {
            return InvokerService.InvokeInternal.ExecuteInt32MethodGet(this, "DrillUp", pivotItem);
        }

        /// <summary>
        /// SupportByVersion Excel 15,16
        /// </summary>
        /// <param name="pivotItem">NetOffice.ExcelApi.PivotItem pivotItem</param>
        /// <param name="pivotLine">optional object pivotLine</param>
        [CustomMethod]
        [SupportByVersion("Excel", 15, 16)]
        public virtual Int32 DrillUp(NetOffice.ExcelApi.PivotItem pivotItem, object pivotLine)
        {
            return InvokerService.InvokeInternal.ExecuteInt32MethodGet(this, "DrillUp", pivotItem, pivotLine);
        }

        /// <summary>
        /// SupportByVersion Excel 15,16
        /// </summary>
        /// <param name="pivotItem">NetOffice.ExcelApi.PivotItem pivotItem</param>
        /// <param name="cubeField">NetOffice.ExcelApi.CubeField cubeField</param>
        /// <param name="pivotLine">optional object pivotLine</param>
        [SupportByVersion("Excel", 15, 16)]
        public virtual Int32 DrillTo(NetOffice.ExcelApi.PivotItem pivotItem, NetOffice.ExcelApi.CubeField cubeField, object pivotLine)
        {
            return InvokerService.InvokeInternal.ExecuteInt32MethodGet(this, "DrillTo", pivotItem, cubeField, pivotLine);
        }

        /// <summary>
        /// SupportByVersion Excel 15,16
        /// </summary>
        /// <param name="pivotItem">NetOffice.ExcelApi.PivotItem pivotItem</param>
        /// <param name="cubeField">NetOffice.ExcelApi.CubeField cubeField</param>
        [CustomMethod]
        [SupportByVersion("Excel", 15, 16)]
        public virtual Int32 DrillTo(NetOffice.ExcelApi.PivotItem pivotItem, NetOffice.ExcelApi.CubeField cubeField)
        {
            return InvokerService.InvokeInternal.ExecuteInt32MethodGet(this, "DrillTo", pivotItem, cubeField);
        }

        /// <summary>
        /// SupportByVersion Excel 15,16
        /// </summary>
        /// <param name="arg1">object arg1</param>
        /// <param name="arg2">optional object arg2</param>
        /// <param name="arg3">optional object arg3</param>
        /// <param name="arg4">optional object arg4</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [SupportByVersion("Excel", 15, 16)]
        public virtual object Dummy2(object arg1, object arg2, object arg3, object arg4)
        {
            return InvokerService.InvokeInternal.ExecuteVariantMethodGet(this, "Dummy2", arg1, arg2, arg3, arg4);
        }

        /// <summary>
        /// SupportByVersion Excel 15,16
        /// </summary>
        /// <param name="arg1">object arg1</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [CustomMethod]
        [SupportByVersion("Excel", 15, 16)]
        public virtual object Dummy2(object arg1)
        {
            return InvokerService.InvokeInternal.ExecuteVariantMethodGet(this, "Dummy2", arg1);
        }

        /// <summary>
        /// SupportByVersion Excel 15,16
        /// </summary>
        /// <param name="arg1">object arg1</param>
        /// <param name="arg2">optional object arg2</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [CustomMethod]
        [SupportByVersion("Excel", 15, 16)]
        public virtual object Dummy2(object arg1, object arg2)
        {
            return InvokerService.InvokeInternal.ExecuteVariantMethodGet(this, "Dummy2", arg1, arg2);
        }

        /// <summary>
        /// SupportByVersion Excel 15,16
        /// </summary>
        /// <param name="arg1">object arg1</param>
        /// <param name="arg2">optional object arg2</param>
        /// <param name="arg3">optional object arg3</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [CustomMethod]
        [SupportByVersion("Excel", 15, 16)]
        public virtual object Dummy2(object arg1, object arg2, object arg3)
        {
            return InvokerService.InvokeInternal.ExecuteVariantMethodGet(this, "Dummy2", arg1, arg2, arg3);
        }

        #endregion

        #pragma warning restore
    }
}

