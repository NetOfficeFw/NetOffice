using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.ExcelApi
{
    /// <summary>
    /// PivotTable
    /// </summary>
    [SyntaxBypass]
    public interface PivotTable_ : ICOMObject
    {
        #region Properties

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// Unknown COM Proxy
        /// </summary>
        /// <param name="index">optional object index</param>
        /// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff839050.aspx
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16), ProxyResult]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        object get_ColumnFields(object index);

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff839050.aspx
        /// Alias for get_ColumnFields
        /// Unknown COM Proxy
        /// </summary>
        /// <param name="index">optional object index</param>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16), ProxyResult, Redirect("get_ColumnFields")]
        object ColumnFields(object index);

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// Unknown COM Proxy
        /// </summary>
        /// <param name="index">optional object index</param>
        /// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff196291.aspx
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16), ProxyResult]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        object get_DataFields(object index);

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff196291.aspx
        /// Alias for get_DataFields
        /// Unknown COM Proxy
        /// </summary>
        /// <param name="index">optional object index</param>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16), ProxyResult, Redirect("get_DataFields")]
        object DataFields(object index);

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// Unknown COM Proxy
        /// </summary>
        /// <param name="index">optional object index</param>
        /// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff841004.aspx
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16), ProxyResult]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        object get_HiddenFields(object index);

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff841004.aspx
        /// Alias for get_HiddenFields
        /// Unknown COM Proxy
        /// </summary>
        /// <param name="index">optional object index</param>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16), ProxyResult, Redirect("get_HiddenFields")]
        object HiddenFields(object index);

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// Unknown COM Proxy
        /// </summary>
        /// <param name="index">optional object index</param>
        /// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff840731.aspx
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16), ProxyResult]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        object get_PageFields(object index);

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff840731.aspx
        /// Alias for get_PageFields
        /// Unknown COM Proxy
        /// </summary>
        /// <param name="index">optional object index</param>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16), ProxyResult, Redirect("get_PageFields")]
        object PageFields(object index);

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// Unknown COM Proxy
        /// </summary>
        /// <param name="index">optional object index</param>
        /// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff196706.aspx
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16), ProxyResult]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        object get_RowFields(object index);

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff196706.aspx
        /// Alias for get_RowFields
        /// Unknown COM Proxy
        /// </summary>
        /// <param name="index">optional object index</param>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16), ProxyResult, Redirect("get_RowFields")]
        object RowFields(object index);

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// Unknown COM Proxy
        /// </summary>
        /// <param name="index">optional object index</param>
        /// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff192982.aspx
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16), ProxyResult]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        object get_VisibleFields(object index);

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff192982.aspx
        /// Alias for get_VisibleFields
        /// Unknown COM Proxy
        /// </summary>
        /// <param name="index">optional object index</param>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16), ProxyResult, Redirect("get_VisibleFields")]
        object VisibleFields(object index);

        #endregion
    }

    /// <summary>
    /// DispatchInterface PivotTable 
    /// SupportByVersion Excel, 9,10,11,12,14,15,16
    /// </summary>
    /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff837611.aspx </remarks>
    [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
    [EntityType(EntityType.IsDispatchInterface)]
    public interface PivotTable : PivotTable_
    {
        #region Properties

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff836434.aspx </remarks>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        NetOffice.ExcelApi.Application Application { get; }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff822808.aspx </remarks>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        NetOffice.ExcelApi.Enums.XlCreator Creator { get; }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// Unknown COM Proxy
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194991.aspx </remarks>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16), ProxyResult]
        object Parent { get; }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// Unknown COM Proxy
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff839050.aspx </remarks>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16), ProxyResult]
        new object ColumnFields { get; }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff837615.aspx </remarks>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        bool ColumnGrand { get; set; }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff834700.aspx </remarks>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        NetOffice.ExcelApi.Range ColumnRange { get; }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff837966.aspx </remarks>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        NetOffice.ExcelApi.Range DataBodyRange { get; }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// Unknown COM Proxy
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff196291.aspx </remarks>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16), ProxyResult]
        new object DataFields { get; }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff836518.aspx </remarks>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        NetOffice.ExcelApi.Range DataLabelRange { get; }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        string _Default { get; set; }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        bool HasAutoFormat { get; set; }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// Unknown COM Proxy
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff841004.aspx </remarks>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16), ProxyResult]
        new object HiddenFields { get; }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff196630.aspx </remarks>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        string InnerDetail { get; set; }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff834372.aspx </remarks>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        string Name { get; set; }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// Unknown COM Proxy
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff840731.aspx </remarks>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16), ProxyResult]
        new object PageFields { get; }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff193268.aspx </remarks>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        NetOffice.ExcelApi.Range PageRange { get; }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194754.aspx </remarks>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        NetOffice.ExcelApi.Range PageRangeCells { get; }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff834610.aspx </remarks>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        DateTime RefreshDate { get; }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff197789.aspx </remarks>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        string RefreshName { get; }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// Unknown COM Proxy
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff196706.aspx </remarks>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16), ProxyResult]
        new object RowFields { get; }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff836789.aspx </remarks>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        bool RowGrand { get; set; }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff196897.aspx </remarks>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        NetOffice.ExcelApi.Range RowRange { get; }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff841136.aspx </remarks>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        bool SaveData { get; set; }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff193521.aspx </remarks>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        object SourceData { get; set; }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff198140.aspx </remarks>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        NetOffice.ExcelApi.Range TableRange1 { get; }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff834378.aspx </remarks>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        NetOffice.ExcelApi.Range TableRange2 { get; }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff837601.aspx </remarks>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        string Value { get; set; }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// Unknown COM Proxy
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff192982.aspx </remarks>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16), ProxyResult]
        new object VisibleFields { get; }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff841243.aspx </remarks>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        Int32 CacheIndex { get; set; }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff821032.aspx </remarks>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        bool DisplayErrorString { get; set; }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff837793.aspx </remarks>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        bool DisplayNullString { get; set; }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff196269.aspx </remarks>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        bool EnableDrilldown { get; set; }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff197903.aspx </remarks>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        bool EnableFieldDialog { get; set; }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff197150.aspx </remarks>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        bool EnableWizard { get; set; }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff834682.aspx </remarks>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        string ErrorString { get; set; }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff823168.aspx </remarks>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        bool ManualUpdate { get; set; }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff195828.aspx </remarks>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        bool MergeLabels { get; set; }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff841149.aspx </remarks>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        string NullString { get; set; }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff841207.aspx </remarks>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        NetOffice.ExcelApi.PivotFormulas PivotFormulas { get; }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838394.aspx </remarks>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        bool SubtotalHiddenPageItems { get; set; }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff193671.aspx </remarks>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        Int32 PageFieldOrder { get; set; }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff835276.aspx </remarks>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        string PageFieldStyle { get; set; }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff836150.aspx </remarks>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        Int32 PageFieldWrapCount { get; set; }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff839462.aspx </remarks>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        bool PreserveFormatting { get; set; }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff840724.aspx </remarks>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        string PivotSelection { get; set; }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff822334.aspx </remarks>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        NetOffice.ExcelApi.Enums.XlPTSelectionMode SelectionMode { get; set; }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        string TableStyle { get; set; }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff834680.aspx </remarks>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        string Tag { get; set; }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff836190.aspx </remarks>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        string VacatedStyle { get; set; }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff837570.aspx </remarks>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        bool PrintTitles { get; set; }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff193066.aspx </remarks>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        NetOffice.ExcelApi.CubeFields CubeFields { get; }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff834419.aspx </remarks>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        string GrandTotalName { get; set; }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff837814.aspx </remarks>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        bool SmallGrid { get; set; }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff836232.aspx </remarks>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        bool RepeatItemsOnEachPrintedPage { get; set; }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff839225.aspx </remarks>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        bool TotalsAnnotation { get; set; }

        /// <summary>
        /// SupportByVersion Excel 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff822897.aspx </remarks>
        [SupportByVersion("Excel", 10, 11, 12, 14, 15, 16)]
        string PivotSelectionStandard { get; set; }

        /// <summary>
        /// SupportByVersion Excel 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff192958.aspx </remarks>
        [SupportByVersion("Excel", 10, 11, 12, 14, 15, 16)]
        NetOffice.ExcelApi.PivotField DataPivotField { get; }

        /// <summary>
        /// SupportByVersion Excel 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff821016.aspx </remarks>
        [SupportByVersion("Excel", 10, 11, 12, 14, 15, 16)]
        bool EnableDataValueEditing { get; set; }

        /// <summary>
        /// SupportByVersion Excel 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff198299.aspx </remarks>
        [SupportByVersion("Excel", 10, 11, 12, 14, 15, 16)]
        string MDX { get; }

        /// <summary>
        /// SupportByVersion Excel 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff195847.aspx </remarks>
        [SupportByVersion("Excel", 10, 11, 12, 14, 15, 16)]
        bool ViewCalculatedMembers { get; set; }

        /// <summary>
        /// SupportByVersion Excel 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff821979.aspx </remarks>
        [SupportByVersion("Excel", 10, 11, 12, 14, 15, 16)]
        NetOffice.ExcelApi.CalculatedMembers CalculatedMembers { get; }

        /// <summary>
        /// SupportByVersion Excel 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff834347.aspx </remarks>
        [SupportByVersion("Excel", 10, 11, 12, 14, 15, 16)]
        bool DisplayImmediateItems { get; set; }

        /// <summary>
        /// SupportByVersion Excel 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff197173.aspx </remarks>
        [SupportByVersion("Excel", 10, 11, 12, 14, 15, 16)]
        bool EnableFieldList { get; set; }

        /// <summary>
        /// SupportByVersion Excel 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff195800.aspx </remarks>
        [SupportByVersion("Excel", 10, 11, 12, 14, 15, 16)]
        bool VisualTotals { get; set; }

        /// <summary>
        /// SupportByVersion Excel 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff196070.aspx </remarks>
        [SupportByVersion("Excel", 10, 11, 12, 14, 15, 16)]
        bool ShowPageMultipleItemLabel { get; set; }

        /// <summary>
        /// SupportByVersion Excel 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff822343.aspx </remarks>
        [SupportByVersion("Excel", 10, 11, 12, 14, 15, 16)]
        NetOffice.ExcelApi.Enums.XlPivotTableVersionList Version { get; }

        /// <summary>
        /// SupportByVersion Excel 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838653.aspx </remarks>
        [SupportByVersion("Excel", 10, 11, 12, 14, 15, 16)]
        bool DisplayEmptyRow { get; set; }

        /// <summary>
        /// SupportByVersion Excel 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff821107.aspx </remarks>
        [SupportByVersion("Excel", 10, 11, 12, 14, 15, 16)]
        bool DisplayEmptyColumn { get; set; }

        /// <summary>
        /// SupportByVersion Excel 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Excel", 10, 11, 12, 14, 15, 16)]
        bool ShowCellBackgroundFromOLAP { get; set; }

        /// <summary>
        /// SupportByVersion Excel 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff193536.aspx </remarks>
        [SupportByVersion("Excel", 12, 14, 15, 16)]
        NetOffice.ExcelApi.PivotAxis PivotColumnAxis { get; }

        /// <summary>
        /// SupportByVersion Excel 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff195054.aspx </remarks>
        [SupportByVersion("Excel", 12, 14, 15, 16)]
        NetOffice.ExcelApi.PivotAxis PivotRowAxis { get; }

        /// <summary>
        /// SupportByVersion Excel 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff823075.aspx </remarks>
        [SupportByVersion("Excel", 12, 14, 15, 16)]
        bool ShowDrillIndicators { get; set; }

        /// <summary>
        /// SupportByVersion Excel 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff839363.aspx </remarks>
        [SupportByVersion("Excel", 12, 14, 15, 16)]
        bool PrintDrillIndicators { get; set; }

        /// <summary>
        /// SupportByVersion Excel 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff839027.aspx </remarks>
        [SupportByVersion("Excel", 12, 14, 15, 16)]
        bool DisplayMemberPropertyTooltips { get; set; }

        /// <summary>
        /// SupportByVersion Excel 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff839074.aspx </remarks>
        [SupportByVersion("Excel", 12, 14, 15, 16)]
        bool DisplayContextTooltips { get; set; }

        /// <summary>
        /// SupportByVersion Excel 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194525.aspx </remarks>
        [SupportByVersion("Excel", 12, 14, 15, 16)]
        Int32 CompactRowIndent { get; set; }

        /// <summary>
        /// SupportByVersion Excel 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff840601.aspx </remarks>
        [SupportByVersion("Excel", 12, 14, 15, 16)]
        NetOffice.ExcelApi.Enums.XlLayoutRowType LayoutRowDefault { get; set; }

        /// <summary>
        /// SupportByVersion Excel 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff837102.aspx </remarks>
        [SupportByVersion("Excel", 12, 14, 15, 16)]
        bool DisplayFieldCaptions { get; set; }

        /// <summary>
        /// SupportByVersion Excel 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff196553.aspx </remarks>
        [SupportByVersion("Excel", 12, 14, 15, 16)]
        NetOffice.ExcelApi.PivotFilters ActiveFilters { get; }

        /// <summary>
        /// SupportByVersion Excel 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff197576.aspx </remarks>
        [SupportByVersion("Excel", 12, 14, 15, 16)]
        bool InGridDropZones { get; set; }

        /// <summary>
        /// SupportByVersion Excel 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff839448.aspx </remarks>
        [SupportByVersion("Excel", 12, 14, 15, 16)]
        object TableStyle2 { get; set; }

        /// <summary>
        /// SupportByVersion Excel 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Excel", 12, 14, 15, 16)]
        bool ShowTableStyleLastColumn { get; set; }

        /// <summary>
        /// SupportByVersion Excel 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff821205.aspx </remarks>
        [SupportByVersion("Excel", 12, 14, 15, 16)]
        bool ShowTableStyleRowStripes { get; set; }

        /// <summary>
        /// SupportByVersion Excel 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff841089.aspx </remarks>
        [SupportByVersion("Excel", 12, 14, 15, 16)]
        bool ShowTableStyleColumnStripes { get; set; }

        /// <summary>
        /// SupportByVersion Excel 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff195083.aspx </remarks>
        [SupportByVersion("Excel", 12, 14, 15, 16)]
        bool ShowTableStyleRowHeaders { get; set; }

        /// <summary>
        /// SupportByVersion Excel 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194144.aspx </remarks>
        [SupportByVersion("Excel", 12, 14, 15, 16)]
        bool ShowTableStyleColumnHeaders { get; set; }

        /// <summary>
        /// SupportByVersion Excel 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff840341.aspx </remarks>
        [SupportByVersion("Excel", 12, 14, 15, 16)]
        bool AllowMultipleFilters { get; set; }

        /// <summary>
        /// SupportByVersion Excel 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff836831.aspx </remarks>
        [SupportByVersion("Excel", 12, 14, 15, 16)]
        string CompactLayoutRowHeader { get; set; }

        /// <summary>
        /// SupportByVersion Excel 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff821896.aspx </remarks>
        [SupportByVersion("Excel", 12, 14, 15, 16)]
        string CompactLayoutColumnHeader { get; set; }

        /// <summary>
        /// SupportByVersion Excel 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff839635.aspx </remarks>
        [SupportByVersion("Excel", 12, 14, 15, 16)]
        bool FieldListSortAscending { get; set; }

        /// <summary>
        /// SupportByVersion Excel 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff841270.aspx </remarks>
        [SupportByVersion("Excel", 12, 14, 15, 16)]
        bool SortUsingCustomLists { get; set; }

        /// <summary>
        /// SupportByVersion Excel 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff820853.aspx </remarks>
        [SupportByVersion("Excel", 12, 14, 15, 16)]
        string Location { get; set; }

        /// <summary>
        /// SupportByVersion Excel 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff839386.aspx </remarks>
        [SupportByVersion("Excel", 14, 15, 16)]
        bool EnableWriteback { get; set; }

        /// <summary>
        /// SupportByVersion Excel 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff837766.aspx </remarks>
        [SupportByVersion("Excel", 14, 15, 16)]
        NetOffice.ExcelApi.Enums.XlAllocation Allocation { get; set; }

        /// <summary>
        /// SupportByVersion Excel 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838849.aspx </remarks>
        [SupportByVersion("Excel", 14, 15, 16)]
        NetOffice.ExcelApi.Enums.XlAllocationValue AllocationValue { get; set; }

        /// <summary>
        /// SupportByVersion Excel 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff822906.aspx </remarks>
        [SupportByVersion("Excel", 14, 15, 16)]
        NetOffice.ExcelApi.Enums.XlAllocationMethod AllocationMethod { get; set; }

        /// <summary>
        /// SupportByVersion Excel 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff836470.aspx </remarks>
        [SupportByVersion("Excel", 14, 15, 16)]
        string AllocationWeightExpression { get; set; }

        /// <summary>
        /// SupportByVersion Excel 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff195057.aspx </remarks>
        [SupportByVersion("Excel", 14, 15, 16)]
        NetOffice.ExcelApi.PivotTableChangeList ChangeList { get; }

        /// <summary>
        /// SupportByVersion Excel 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff839681.aspx </remarks>
        [SupportByVersion("Excel", 14, 15, 16)]
        NetOffice.ExcelApi.Slicers Slicers { get; }

        /// <summary>
        /// SupportByVersion Excel 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838986.aspx </remarks>
        [SupportByVersion("Excel", 14, 15, 16)]
        string AlternativeText { get; set; }

        /// <summary>
        /// SupportByVersion Excel 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff198197.aspx </remarks>
        [SupportByVersion("Excel", 14, 15, 16)]
        string Summary { get; set; }

        /// <summary>
        /// SupportByVersion Excel 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838806.aspx </remarks>
        [SupportByVersion("Excel", 14, 15, 16)]
        bool VisualTotalsForSets { get; set; }

        /// <summary>
        /// SupportByVersion Excel 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff835567.aspx </remarks>
        [SupportByVersion("Excel", 14, 15, 16)]
        bool ShowValuesRow { get; set; }

        /// <summary>
        /// SupportByVersion Excel 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194933.aspx </remarks>
        [SupportByVersion("Excel", 14, 15, 16)]
        bool CalculatedMembersInFilters { get; set; }

        /// <summary>
        /// SupportByVersion Excel 15,16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/jj231466.aspx </remarks>
        [SupportByVersion("Excel", 15, 16)]
        bool Hidden { get; }

        /// <summary>
        /// SupportByVersion Excel 15,16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/jj227930.aspx </remarks>
        [SupportByVersion("Excel", 15, 16)]
        NetOffice.ExcelApi.Shape PivotChart { get; }

        #endregion

        #region Methods

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff837987.aspx </remarks>
        /// <param name="rowFields">optional object rowFields</param>
        /// <param name="columnFields">optional object columnFields</param>
        /// <param name="pageFields">optional object pageFields</param>
        /// <param name="addToTable">optional object addToTable</param>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        object AddFields(object rowFields, object columnFields, object pageFields, object addToTable);

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff837987.aspx </remarks>
        [CustomMethod]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        object AddFields();

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff837987.aspx </remarks>
        /// <param name="rowFields">optional object rowFields</param>
        [CustomMethod]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        object AddFields(object rowFields);

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff837987.aspx </remarks>
        /// <param name="rowFields">optional object rowFields</param>
        /// <param name="columnFields">optional object columnFields</param>
        [CustomMethod]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        object AddFields(object rowFields, object columnFields);

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff837987.aspx </remarks>
        /// <param name="rowFields">optional object rowFields</param>
        /// <param name="columnFields">optional object columnFields</param>
        /// <param name="pageFields">optional object pageFields</param>
        [CustomMethod]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        object AddFields(object rowFields, object columnFields, object pageFields);

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff834670.aspx </remarks>
        /// <param name="pageField">optional object pageField</param>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        object ShowPages(object pageField);

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff834670.aspx </remarks>
        [CustomMethod]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        object ShowPages();

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff195453.aspx </remarks>
        /// <param name="index">optional object index</param>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        object PivotFields(object index);

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff195453.aspx </remarks>
        [CustomMethod]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        object PivotFields();

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff834300.aspx </remarks>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        bool RefreshTable();

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff835843.aspx </remarks>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        NetOffice.ExcelApi.CalculatedFields CalculatedFields();

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838792.aspx </remarks>
        /// <param name="name">string name</param>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        Double GetData(string name);

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff197802.aspx </remarks>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        void ListFormulas();

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff834938.aspx </remarks>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        NetOffice.ExcelApi.PivotCache PivotCache();

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff821973.aspx </remarks>
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
        void PivotTableWizard(object sourceType, object sourceData, object tableDestination, object tableName, object rowGrand, object columnGrand, object saveData, object hasAutoFormat, object autoPage, object reserved, object backgroundQuery, object optimizeCache, object pageFieldOrder, object pageFieldWrapCount, object readData, object connection);

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff821973.aspx </remarks>
        [CustomMethod]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        void PivotTableWizard();

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff821973.aspx </remarks>
        /// <param name="sourceType">optional object sourceType</param>
        [CustomMethod]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        void PivotTableWizard(object sourceType);

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff821973.aspx </remarks>
        /// <param name="sourceType">optional object sourceType</param>
        /// <param name="sourceData">optional object sourceData</param>
        [CustomMethod]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        void PivotTableWizard(object sourceType, object sourceData);

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff821973.aspx </remarks>
        /// <param name="sourceType">optional object sourceType</param>
        /// <param name="sourceData">optional object sourceData</param>
        /// <param name="tableDestination">optional object tableDestination</param>
        [CustomMethod]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        void PivotTableWizard(object sourceType, object sourceData, object tableDestination);

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff821973.aspx </remarks>
        /// <param name="sourceType">optional object sourceType</param>
        /// <param name="sourceData">optional object sourceData</param>
        /// <param name="tableDestination">optional object tableDestination</param>
        /// <param name="tableName">optional object tableName</param>
        [CustomMethod]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        void PivotTableWizard(object sourceType, object sourceData, object tableDestination, object tableName);

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff821973.aspx </remarks>
        /// <param name="sourceType">optional object sourceType</param>
        /// <param name="sourceData">optional object sourceData</param>
        /// <param name="tableDestination">optional object tableDestination</param>
        /// <param name="tableName">optional object tableName</param>
        /// <param name="rowGrand">optional object rowGrand</param>
        [CustomMethod]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        void PivotTableWizard(object sourceType, object sourceData, object tableDestination, object tableName, object rowGrand);

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff821973.aspx </remarks>
        /// <param name="sourceType">optional object sourceType</param>
        /// <param name="sourceData">optional object sourceData</param>
        /// <param name="tableDestination">optional object tableDestination</param>
        /// <param name="tableName">optional object tableName</param>
        /// <param name="rowGrand">optional object rowGrand</param>
        /// <param name="columnGrand">optional object columnGrand</param>
        [CustomMethod]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        void PivotTableWizard(object sourceType, object sourceData, object tableDestination, object tableName, object rowGrand, object columnGrand);

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff821973.aspx </remarks>
        /// <param name="sourceType">optional object sourceType</param>
        /// <param name="sourceData">optional object sourceData</param>
        /// <param name="tableDestination">optional object tableDestination</param>
        /// <param name="tableName">optional object tableName</param>
        /// <param name="rowGrand">optional object rowGrand</param>
        /// <param name="columnGrand">optional object columnGrand</param>
        /// <param name="saveData">optional object saveData</param>
        [CustomMethod]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        void PivotTableWizard(object sourceType, object sourceData, object tableDestination, object tableName, object rowGrand, object columnGrand, object saveData);

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff821973.aspx </remarks>
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
        void PivotTableWizard(object sourceType, object sourceData, object tableDestination, object tableName, object rowGrand, object columnGrand, object saveData, object hasAutoFormat);

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff821973.aspx </remarks>
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
        void PivotTableWizard(object sourceType, object sourceData, object tableDestination, object tableName, object rowGrand, object columnGrand, object saveData, object hasAutoFormat, object autoPage);

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff821973.aspx </remarks>
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
        void PivotTableWizard(object sourceType, object sourceData, object tableDestination, object tableName, object rowGrand, object columnGrand, object saveData, object hasAutoFormat, object autoPage, object reserved);

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff821973.aspx </remarks>
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
        void PivotTableWizard(object sourceType, object sourceData, object tableDestination, object tableName, object rowGrand, object columnGrand, object saveData, object hasAutoFormat, object autoPage, object reserved, object backgroundQuery);

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff821973.aspx </remarks>
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
        void PivotTableWizard(object sourceType, object sourceData, object tableDestination, object tableName, object rowGrand, object columnGrand, object saveData, object hasAutoFormat, object autoPage, object reserved, object backgroundQuery, object optimizeCache);

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff821973.aspx </remarks>
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
        void PivotTableWizard(object sourceType, object sourceData, object tableDestination, object tableName, object rowGrand, object columnGrand, object saveData, object hasAutoFormat, object autoPage, object reserved, object backgroundQuery, object optimizeCache, object pageFieldOrder);

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff821973.aspx </remarks>
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
        void PivotTableWizard(object sourceType, object sourceData, object tableDestination, object tableName, object rowGrand, object columnGrand, object saveData, object hasAutoFormat, object autoPage, object reserved, object backgroundQuery, object optimizeCache, object pageFieldOrder, object pageFieldWrapCount);

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff821973.aspx </remarks>
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
        void PivotTableWizard(object sourceType, object sourceData, object tableDestination, object tableName, object rowGrand, object columnGrand, object saveData, object hasAutoFormat, object autoPage, object reserved, object backgroundQuery, object optimizeCache, object pageFieldOrder, object pageFieldWrapCount, object readData);

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff840451.aspx </remarks>
        /// <param name="name">string name</param>
        /// <param name="mode">optional NetOffice.ExcelApi.Enums.XlPTSelectionMode Mode = 0</param>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        void PivotSelect(string name, object mode);

        /// <summary>
        /// SupportByVersion Excel 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff840451.aspx </remarks>
        /// <param name="name">string name</param>
        /// <param name="mode">optional NetOffice.ExcelApi.Enums.XlPTSelectionMode Mode = 0</param>
        /// <param name="useStandardName">optional object useStandardName</param>
        [SupportByVersion("Excel", 10, 11, 12, 14, 15, 16)]
        void PivotSelect(string name, object mode, object useStandardName);

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff840451.aspx </remarks>
        /// <param name="name">string name</param>
        [CustomMethod]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        void PivotSelect(string name);

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff196581.aspx </remarks>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        void Update();

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="format">NetOffice.ExcelApi.Enums.xlPivotFormatType format</param>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        void Format(NetOffice.ExcelApi.Enums.xlPivotFormatType format);

        /// <summary>
        /// SupportByVersion Excel 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="name">string name</param>
        /// <param name="mode">optional NetOffice.ExcelApi.Enums.XlPTSelectionMode Mode = 0</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [SupportByVersion("Excel", 10, 11, 12, 14, 15, 16)]
        void _PivotSelect(string name, object mode);

        /// <summary>
        /// SupportByVersion Excel 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="name">string name</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [CustomMethod]
        [SupportByVersion("Excel", 10, 11, 12, 14, 15, 16)]
        void _PivotSelect(string name);

        /// <summary>
        /// SupportByVersion Excel 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff195919.aspx </remarks>
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
        NetOffice.ExcelApi.Range GetPivotData(object dataField, object field1, object item1, object field2, object item2, object field3, object item3, object field4, object item4, object field5, object item5, object field6, object item6, object field7, object item7, object field8, object item8, object field9, object item9, object field10, object item10, object field11, object item11, object field12, object item12, object field13, object item13, object field14, object item14);

        /// <summary>
        /// SupportByVersion Excel 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff195919.aspx </remarks>
        [CustomMethod]
        [SupportByVersion("Excel", 10, 11, 12, 14, 15, 16)]
        NetOffice.ExcelApi.Range GetPivotData();

        /// <summary>
        /// SupportByVersion Excel 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff195919.aspx </remarks>
        /// <param name="dataField">optional object dataField</param>
        [CustomMethod]
        [SupportByVersion("Excel", 10, 11, 12, 14, 15, 16)]
        NetOffice.ExcelApi.Range GetPivotData(object dataField);

        /// <summary>
        /// SupportByVersion Excel 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff195919.aspx </remarks>
        /// <param name="dataField">optional object dataField</param>
        /// <param name="field1">optional object field1</param>
        [CustomMethod]
        [SupportByVersion("Excel", 10, 11, 12, 14, 15, 16)]
        NetOffice.ExcelApi.Range GetPivotData(object dataField, object field1);

        /// <summary>
        /// SupportByVersion Excel 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff195919.aspx </remarks>
        /// <param name="dataField">optional object dataField</param>
        /// <param name="field1">optional object field1</param>
        /// <param name="item1">optional object item1</param>
        [CustomMethod]
        [SupportByVersion("Excel", 10, 11, 12, 14, 15, 16)]
        NetOffice.ExcelApi.Range GetPivotData(object dataField, object field1, object item1);

        /// <summary>
        /// SupportByVersion Excel 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff195919.aspx </remarks>
        /// <param name="dataField">optional object dataField</param>
        /// <param name="field1">optional object field1</param>
        /// <param name="item1">optional object item1</param>
        /// <param name="field2">optional object field2</param>
        [CustomMethod]
        [SupportByVersion("Excel", 10, 11, 12, 14, 15, 16)]
        NetOffice.ExcelApi.Range GetPivotData(object dataField, object field1, object item1, object field2);

        /// <summary>
        /// SupportByVersion Excel 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff195919.aspx </remarks>
        /// <param name="dataField">optional object dataField</param>
        /// <param name="field1">optional object field1</param>
        /// <param name="item1">optional object item1</param>
        /// <param name="field2">optional object field2</param>
        /// <param name="item2">optional object item2</param>
        [CustomMethod]
        [SupportByVersion("Excel", 10, 11, 12, 14, 15, 16)]
        NetOffice.ExcelApi.Range GetPivotData(object dataField, object field1, object item1, object field2, object item2);

        /// <summary>
        /// SupportByVersion Excel 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff195919.aspx </remarks>
        /// <param name="dataField">optional object dataField</param>
        /// <param name="field1">optional object field1</param>
        /// <param name="item1">optional object item1</param>
        /// <param name="field2">optional object field2</param>
        /// <param name="item2">optional object item2</param>
        /// <param name="field3">optional object field3</param>
        [CustomMethod]
        [SupportByVersion("Excel", 10, 11, 12, 14, 15, 16)]
        NetOffice.ExcelApi.Range GetPivotData(object dataField, object field1, object item1, object field2, object item2, object field3);

        /// <summary>
        /// SupportByVersion Excel 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff195919.aspx </remarks>
        /// <param name="dataField">optional object dataField</param>
        /// <param name="field1">optional object field1</param>
        /// <param name="item1">optional object item1</param>
        /// <param name="field2">optional object field2</param>
        /// <param name="item2">optional object item2</param>
        /// <param name="field3">optional object field3</param>
        /// <param name="item3">optional object item3</param>
        [CustomMethod]
        [SupportByVersion("Excel", 10, 11, 12, 14, 15, 16)]
        NetOffice.ExcelApi.Range GetPivotData(object dataField, object field1, object item1, object field2, object item2, object field3, object item3);

        /// <summary>
        /// SupportByVersion Excel 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff195919.aspx </remarks>
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
        NetOffice.ExcelApi.Range GetPivotData(object dataField, object field1, object item1, object field2, object item2, object field3, object item3, object field4);

        /// <summary>
        /// SupportByVersion Excel 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff195919.aspx </remarks>
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
        NetOffice.ExcelApi.Range GetPivotData(object dataField, object field1, object item1, object field2, object item2, object field3, object item3, object field4, object item4);

        /// <summary>
        /// SupportByVersion Excel 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff195919.aspx </remarks>
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
        NetOffice.ExcelApi.Range GetPivotData(object dataField, object field1, object item1, object field2, object item2, object field3, object item3, object field4, object item4, object field5);

        /// <summary>
        /// SupportByVersion Excel 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff195919.aspx </remarks>
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
        NetOffice.ExcelApi.Range GetPivotData(object dataField, object field1, object item1, object field2, object item2, object field3, object item3, object field4, object item4, object field5, object item5);

        /// <summary>
        /// SupportByVersion Excel 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff195919.aspx </remarks>
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
        NetOffice.ExcelApi.Range GetPivotData(object dataField, object field1, object item1, object field2, object item2, object field3, object item3, object field4, object item4, object field5, object item5, object field6);

        /// <summary>
        /// SupportByVersion Excel 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff195919.aspx </remarks>
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
        NetOffice.ExcelApi.Range GetPivotData(object dataField, object field1, object item1, object field2, object item2, object field3, object item3, object field4, object item4, object field5, object item5, object field6, object item6);

        /// <summary>
        /// SupportByVersion Excel 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff195919.aspx </remarks>
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
        NetOffice.ExcelApi.Range GetPivotData(object dataField, object field1, object item1, object field2, object item2, object field3, object item3, object field4, object item4, object field5, object item5, object field6, object item6, object field7);

        /// <summary>
        /// SupportByVersion Excel 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff195919.aspx </remarks>
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
        NetOffice.ExcelApi.Range GetPivotData(object dataField, object field1, object item1, object field2, object item2, object field3, object item3, object field4, object item4, object field5, object item5, object field6, object item6, object field7, object item7);

        /// <summary>
        /// SupportByVersion Excel 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff195919.aspx </remarks>
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
        NetOffice.ExcelApi.Range GetPivotData(object dataField, object field1, object item1, object field2, object item2, object field3, object item3, object field4, object item4, object field5, object item5, object field6, object item6, object field7, object item7, object field8);

        /// <summary>
        /// SupportByVersion Excel 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff195919.aspx </remarks>
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
        NetOffice.ExcelApi.Range GetPivotData(object dataField, object field1, object item1, object field2, object item2, object field3, object item3, object field4, object item4, object field5, object item5, object field6, object item6, object field7, object item7, object field8, object item8);

        /// <summary>
        /// SupportByVersion Excel 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff195919.aspx </remarks>
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
        NetOffice.ExcelApi.Range GetPivotData(object dataField, object field1, object item1, object field2, object item2, object field3, object item3, object field4, object item4, object field5, object item5, object field6, object item6, object field7, object item7, object field8, object item8, object field9);

        /// <summary>
        /// SupportByVersion Excel 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff195919.aspx </remarks>
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
        NetOffice.ExcelApi.Range GetPivotData(object dataField, object field1, object item1, object field2, object item2, object field3, object item3, object field4, object item4, object field5, object item5, object field6, object item6, object field7, object item7, object field8, object item8, object field9, object item9);

        /// <summary>
        /// SupportByVersion Excel 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff195919.aspx </remarks>
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
        NetOffice.ExcelApi.Range GetPivotData(object dataField, object field1, object item1, object field2, object item2, object field3, object item3, object field4, object item4, object field5, object item5, object field6, object item6, object field7, object item7, object field8, object item8, object field9, object item9, object field10);

        /// <summary>
        /// SupportByVersion Excel 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff195919.aspx </remarks>
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
        NetOffice.ExcelApi.Range GetPivotData(object dataField, object field1, object item1, object field2, object item2, object field3, object item3, object field4, object item4, object field5, object item5, object field6, object item6, object field7, object item7, object field8, object item8, object field9, object item9, object field10, object item10);

        /// <summary>
        /// SupportByVersion Excel 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff195919.aspx </remarks>
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
        NetOffice.ExcelApi.Range GetPivotData(object dataField, object field1, object item1, object field2, object item2, object field3, object item3, object field4, object item4, object field5, object item5, object field6, object item6, object field7, object item7, object field8, object item8, object field9, object item9, object field10, object item10, object field11);

        /// <summary>
        /// SupportByVersion Excel 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff195919.aspx </remarks>
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
        NetOffice.ExcelApi.Range GetPivotData(object dataField, object field1, object item1, object field2, object item2, object field3, object item3, object field4, object item4, object field5, object item5, object field6, object item6, object field7, object item7, object field8, object item8, object field9, object item9, object field10, object item10, object field11, object item11);

        /// <summary>
        /// SupportByVersion Excel 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff195919.aspx </remarks>
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
        NetOffice.ExcelApi.Range GetPivotData(object dataField, object field1, object item1, object field2, object item2, object field3, object item3, object field4, object item4, object field5, object item5, object field6, object item6, object field7, object item7, object field8, object item8, object field9, object item9, object field10, object item10, object field11, object item11, object field12);

        /// <summary>
        /// SupportByVersion Excel 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff195919.aspx </remarks>
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
        NetOffice.ExcelApi.Range GetPivotData(object dataField, object field1, object item1, object field2, object item2, object field3, object item3, object field4, object item4, object field5, object item5, object field6, object item6, object field7, object item7, object field8, object item8, object field9, object item9, object field10, object item10, object field11, object item11, object field12, object item12);

        /// <summary>
        /// SupportByVersion Excel 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff195919.aspx </remarks>
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
        NetOffice.ExcelApi.Range GetPivotData(object dataField, object field1, object item1, object field2, object item2, object field3, object item3, object field4, object item4, object field5, object item5, object field6, object item6, object field7, object item7, object field8, object item8, object field9, object item9, object field10, object item10, object field11, object item11, object field12, object item12, object field13);

        /// <summary>
        /// SupportByVersion Excel 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff195919.aspx </remarks>
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
        NetOffice.ExcelApi.Range GetPivotData(object dataField, object field1, object item1, object field2, object item2, object field3, object item3, object field4, object item4, object field5, object item5, object field6, object item6, object field7, object item7, object field8, object item8, object field9, object item9, object field10, object item10, object field11, object item11, object field12, object item12, object field13, object item13);

        /// <summary>
        /// SupportByVersion Excel 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff195919.aspx </remarks>
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
        NetOffice.ExcelApi.Range GetPivotData(object dataField, object field1, object item1, object field2, object item2, object field3, object item3, object field4, object item4, object field5, object item5, object field6, object item6, object field7, object item7, object field8, object item8, object field9, object item9, object field10, object item10, object field11, object item11, object field12, object item12, object field13, object item13, object field14);

        /// <summary>
        /// SupportByVersion Excel 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff823171.aspx </remarks>
        /// <param name="field">object field</param>
        /// <param name="caption">optional object caption</param>
        /// <param name="function">optional object function</param>
        [SupportByVersion("Excel", 10, 11, 12, 14, 15, 16)]
        NetOffice.ExcelApi.PivotField AddDataField(object field, object caption, object function);

        /// <summary>
        /// SupportByVersion Excel 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff823171.aspx </remarks>
        /// <param name="field">object field</param>
        [CustomMethod]
        [SupportByVersion("Excel", 10, 11, 12, 14, 15, 16)]
        NetOffice.ExcelApi.PivotField AddDataField(object field);

        /// <summary>
        /// SupportByVersion Excel 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff823171.aspx </remarks>
        /// <param name="field">object field</param>
        /// <param name="caption">optional object caption</param>
        [CustomMethod]
        [SupportByVersion("Excel", 10, 11, 12, 14, 15, 16)]
        NetOffice.ExcelApi.PivotField AddDataField(object field, object caption);

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
        object Dummy15(object arg1, object arg2, object arg3, object arg4, object arg5, object arg6, object arg7, object arg8, object arg9, object arg10, object arg11, object arg12, object arg13, object arg14, object arg15, object arg16, object arg17, object arg18, object arg19, object arg20, object arg21, object arg22, object arg23, object arg24, object arg25, object arg26, object arg27, object arg28, object arg29, object arg30);

        /// <summary>
        /// SupportByVersion Excel 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="arg1">object arg1</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [CustomMethod]
        [SupportByVersion("Excel", 10, 11, 12, 14, 15, 16)]
        object Dummy15(object arg1);

        /// <summary>
        /// SupportByVersion Excel 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="arg1">object arg1</param>
        /// <param name="arg2">optional object arg2</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [CustomMethod]
        [SupportByVersion("Excel", 10, 11, 12, 14, 15, 16)]
        object Dummy15(object arg1, object arg2);

        /// <summary>
        /// SupportByVersion Excel 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="arg1">object arg1</param>
        /// <param name="arg2">optional object arg2</param>
        /// <param name="arg3">optional object arg3</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [CustomMethod]
        [SupportByVersion("Excel", 10, 11, 12, 14, 15, 16)]
        object Dummy15(object arg1, object arg2, object arg3);

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
        object Dummy15(object arg1, object arg2, object arg3, object arg4);

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
        object Dummy15(object arg1, object arg2, object arg3, object arg4, object arg5);

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
        object Dummy15(object arg1, object arg2, object arg3, object arg4, object arg5, object arg6);

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
        object Dummy15(object arg1, object arg2, object arg3, object arg4, object arg5, object arg6, object arg7);

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
        object Dummy15(object arg1, object arg2, object arg3, object arg4, object arg5, object arg6, object arg7, object arg8);

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
        object Dummy15(object arg1, object arg2, object arg3, object arg4, object arg5, object arg6, object arg7, object arg8, object arg9);

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
        object Dummy15(object arg1, object arg2, object arg3, object arg4, object arg5, object arg6, object arg7, object arg8, object arg9, object arg10);

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
        object Dummy15(object arg1, object arg2, object arg3, object arg4, object arg5, object arg6, object arg7, object arg8, object arg9, object arg10, object arg11);

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
        object Dummy15(object arg1, object arg2, object arg3, object arg4, object arg5, object arg6, object arg7, object arg8, object arg9, object arg10, object arg11, object arg12);

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
        object Dummy15(object arg1, object arg2, object arg3, object arg4, object arg5, object arg6, object arg7, object arg8, object arg9, object arg10, object arg11, object arg12, object arg13);

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
        object Dummy15(object arg1, object arg2, object arg3, object arg4, object arg5, object arg6, object arg7, object arg8, object arg9, object arg10, object arg11, object arg12, object arg13, object arg14);

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
        object Dummy15(object arg1, object arg2, object arg3, object arg4, object arg5, object arg6, object arg7, object arg8, object arg9, object arg10, object arg11, object arg12, object arg13, object arg14, object arg15);

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
        object Dummy15(object arg1, object arg2, object arg3, object arg4, object arg5, object arg6, object arg7, object arg8, object arg9, object arg10, object arg11, object arg12, object arg13, object arg14, object arg15, object arg16);

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
        object Dummy15(object arg1, object arg2, object arg3, object arg4, object arg5, object arg6, object arg7, object arg8, object arg9, object arg10, object arg11, object arg12, object arg13, object arg14, object arg15, object arg16, object arg17);

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
        object Dummy15(object arg1, object arg2, object arg3, object arg4, object arg5, object arg6, object arg7, object arg8, object arg9, object arg10, object arg11, object arg12, object arg13, object arg14, object arg15, object arg16, object arg17, object arg18);

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
        object Dummy15(object arg1, object arg2, object arg3, object arg4, object arg5, object arg6, object arg7, object arg8, object arg9, object arg10, object arg11, object arg12, object arg13, object arg14, object arg15, object arg16, object arg17, object arg18, object arg19);

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
        object Dummy15(object arg1, object arg2, object arg3, object arg4, object arg5, object arg6, object arg7, object arg8, object arg9, object arg10, object arg11, object arg12, object arg13, object arg14, object arg15, object arg16, object arg17, object arg18, object arg19, object arg20);

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
        object Dummy15(object arg1, object arg2, object arg3, object arg4, object arg5, object arg6, object arg7, object arg8, object arg9, object arg10, object arg11, object arg12, object arg13, object arg14, object arg15, object arg16, object arg17, object arg18, object arg19, object arg20, object arg21);

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
        object Dummy15(object arg1, object arg2, object arg3, object arg4, object arg5, object arg6, object arg7, object arg8, object arg9, object arg10, object arg11, object arg12, object arg13, object arg14, object arg15, object arg16, object arg17, object arg18, object arg19, object arg20, object arg21, object arg22);

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
        object Dummy15(object arg1, object arg2, object arg3, object arg4, object arg5, object arg6, object arg7, object arg8, object arg9, object arg10, object arg11, object arg12, object arg13, object arg14, object arg15, object arg16, object arg17, object arg18, object arg19, object arg20, object arg21, object arg22, object arg23);

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
        object Dummy15(object arg1, object arg2, object arg3, object arg4, object arg5, object arg6, object arg7, object arg8, object arg9, object arg10, object arg11, object arg12, object arg13, object arg14, object arg15, object arg16, object arg17, object arg18, object arg19, object arg20, object arg21, object arg22, object arg23, object arg24);

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
        object Dummy15(object arg1, object arg2, object arg3, object arg4, object arg5, object arg6, object arg7, object arg8, object arg9, object arg10, object arg11, object arg12, object arg13, object arg14, object arg15, object arg16, object arg17, object arg18, object arg19, object arg20, object arg21, object arg22, object arg23, object arg24, object arg25);

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
        object Dummy15(object arg1, object arg2, object arg3, object arg4, object arg5, object arg6, object arg7, object arg8, object arg9, object arg10, object arg11, object arg12, object arg13, object arg14, object arg15, object arg16, object arg17, object arg18, object arg19, object arg20, object arg21, object arg22, object arg23, object arg24, object arg25, object arg26);

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
        object Dummy15(object arg1, object arg2, object arg3, object arg4, object arg5, object arg6, object arg7, object arg8, object arg9, object arg10, object arg11, object arg12, object arg13, object arg14, object arg15, object arg16, object arg17, object arg18, object arg19, object arg20, object arg21, object arg22, object arg23, object arg24, object arg25, object arg26, object arg27);

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
        object Dummy15(object arg1, object arg2, object arg3, object arg4, object arg5, object arg6, object arg7, object arg8, object arg9, object arg10, object arg11, object arg12, object arg13, object arg14, object arg15, object arg16, object arg17, object arg18, object arg19, object arg20, object arg21, object arg22, object arg23, object arg24, object arg25, object arg26, object arg27, object arg28);

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
        object Dummy15(object arg1, object arg2, object arg3, object arg4, object arg5, object arg6, object arg7, object arg8, object arg9, object arg10, object arg11, object arg12, object arg13, object arg14, object arg15, object arg16, object arg17, object arg18, object arg19, object arg20, object arg21, object arg22, object arg23, object arg24, object arg25, object arg26, object arg27, object arg28, object arg29);

        /// <summary>
        /// SupportByVersion Excel 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff821072.aspx </remarks>
        /// <param name="file">string file</param>
        /// <param name="measures">optional object measures</param>
        /// <param name="levels">optional object levels</param>
        /// <param name="members">optional object members</param>
        /// <param name="properties">optional object properties</param>
        [SupportByVersion("Excel", 10, 11, 12, 14, 15, 16)]
        string CreateCubeFile(string file, object measures, object levels, object members, object properties);

        /// <summary>
        /// SupportByVersion Excel 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff821072.aspx </remarks>
        /// <param name="file">string file</param>
        [CustomMethod]
        [SupportByVersion("Excel", 10, 11, 12, 14, 15, 16)]
        string CreateCubeFile(string file);

        /// <summary>
        /// SupportByVersion Excel 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff821072.aspx </remarks>
        /// <param name="file">string file</param>
        /// <param name="measures">optional object measures</param>
        [CustomMethod]
        [SupportByVersion("Excel", 10, 11, 12, 14, 15, 16)]
        string CreateCubeFile(string file, object measures);

        /// <summary>
        /// SupportByVersion Excel 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff821072.aspx </remarks>
        /// <param name="file">string file</param>
        /// <param name="measures">optional object measures</param>
        /// <param name="levels">optional object levels</param>
        [CustomMethod]
        [SupportByVersion("Excel", 10, 11, 12, 14, 15, 16)]
        string CreateCubeFile(string file, object measures, object levels);

        /// <summary>
        /// SupportByVersion Excel 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff821072.aspx </remarks>
        /// <param name="file">string file</param>
        /// <param name="measures">optional object measures</param>
        /// <param name="levels">optional object levels</param>
        /// <param name="members">optional object members</param>
        [CustomMethod]
        [SupportByVersion("Excel", 10, 11, 12, 14, 15, 16)]
        string CreateCubeFile(string file, object measures, object levels, object members);

        /// <summary>
        /// SupportByVersion Excel 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194097.aspx </remarks>
        [SupportByVersion("Excel", 12, 14, 15, 16)]
        void ClearTable();

        /// <summary>
        /// SupportByVersion Excel 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff197262.aspx </remarks>
        /// <param name="rowLayout">NetOffice.ExcelApi.Enums.XlLayoutRowType rowLayout</param>
        [SupportByVersion("Excel", 12, 14, 15, 16)]
        void RowAxisLayout(NetOffice.ExcelApi.Enums.XlLayoutRowType rowLayout);

        /// <summary>
        /// SupportByVersion Excel 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff840038.aspx </remarks>
        /// <param name="location">NetOffice.ExcelApi.Enums.xLSubtototalLocationType location</param>
        [SupportByVersion("Excel", 12, 14, 15, 16)]
        void SubtotalLocation(NetOffice.ExcelApi.Enums.xLSubtototalLocationType location);

        /// <summary>
        /// SupportByVersion Excel 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff840098.aspx </remarks>
        [SupportByVersion("Excel", 12, 14, 15, 16)]
        void ClearAllFilters();

        /// <summary>
        /// SupportByVersion Excel 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff835232.aspx </remarks>
        /// <param name="convertFilters">bool convertFilters</param>
        [SupportByVersion("Excel", 12, 14, 15, 16)]
        void ConvertToFormulas(bool convertFilters);

        /// <summary>
        /// SupportByVersion Excel 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194492.aspx </remarks>
        /// <param name="conn">NetOffice.ExcelApi.WorkbookConnection conn</param>
        [SupportByVersion("Excel", 12, 14, 15, 16)]
        void ChangeConnection(NetOffice.ExcelApi.WorkbookConnection conn);

        /// <summary>
        /// SupportByVersion Excel 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194688.aspx </remarks>
        /// <param name="pivotCache">object pivotCache</param>
        [SupportByVersion("Excel", 12, 14, 15, 16)]
        void ChangePivotCache(object pivotCache);

        /// <summary>
        /// SupportByVersion Excel 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff822662.aspx </remarks>
        [SupportByVersion("Excel", 14, 15, 16)]
        void AllocateChanges();

        /// <summary>
        /// SupportByVersion Excel 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff841032.aspx </remarks>
        [SupportByVersion("Excel", 14, 15, 16)]
        void CommitChanges();

        /// <summary>
        /// SupportByVersion Excel 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff837043.aspx </remarks>
        [SupportByVersion("Excel", 14, 15, 16)]
        void DiscardChanges();

        /// <summary>
        /// SupportByVersion Excel 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff197450.aspx </remarks>
        [SupportByVersion("Excel", 14, 15, 16)]
        void RefreshDataSourceValues();

        /// <summary>
        /// SupportByVersion Excel 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff198076.aspx </remarks>
        /// <param name="repeat">NetOffice.ExcelApi.Enums.XlPivotFieldRepeatLabels repeat</param>
        [SupportByVersion("Excel", 14, 15, 16)]
        void RepeatAllLabels(NetOffice.ExcelApi.Enums.XlPivotFieldRepeatLabels repeat);

        /// <summary>
        /// SupportByVersion Excel 15,16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/jj230950.aspx </remarks>
        /// <param name="rowline">optional object rowline</param>
        /// <param name="columnline">optional object columnline</param>
        [SupportByVersion("Excel", 15, 16)]
        NetOffice.ExcelApi.PivotValueCell PivotValueCell(object rowline, object columnline);

        /// <summary>
        /// SupportByVersion Excel 15,16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/jj230950.aspx </remarks>
        [CustomMethod]
        [SupportByVersion("Excel", 15, 16)]
        NetOffice.ExcelApi.PivotValueCell PivotValueCell();

        /// <summary>
        /// SupportByVersion Excel 15,16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/jj230950.aspx </remarks>
        /// <param name="rowline">optional object rowline</param>
        [CustomMethod]
        [SupportByVersion("Excel", 15, 16)]
        NetOffice.ExcelApi.PivotValueCell PivotValueCell(object rowline);

        /// <summary>
        /// SupportByVersion Excel 15,16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/jj227250.aspx </remarks>
        /// <param name="pivotItem">NetOffice.ExcelApi.PivotItem pivotItem</param>
        /// <param name="pivotLine">optional object pivotLine</param>
        [SupportByVersion("Excel", 15, 16)]
        void DrillDown(NetOffice.ExcelApi.PivotItem pivotItem, object pivotLine);

        /// <summary>
        /// SupportByVersion Excel 15,16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/jj227250.aspx </remarks>
        /// <param name="pivotItem">NetOffice.ExcelApi.PivotItem pivotItem</param>
        [CustomMethod]
        [SupportByVersion("Excel", 15, 16)]
        void DrillDown(NetOffice.ExcelApi.PivotItem pivotItem);

        /// <summary>
        /// SupportByVersion Excel 15,16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/jj227808.aspx </remarks>
        /// <param name="pivotItem">NetOffice.ExcelApi.PivotItem pivotItem</param>
        /// <param name="pivotLine">optional object pivotLine</param>
        /// <param name="levelUniqueName">optional object levelUniqueName</param>
        [SupportByVersion("Excel", 15, 16)]
        void DrillUp(NetOffice.ExcelApi.PivotItem pivotItem, object pivotLine, object levelUniqueName);

        /// <summary>
        /// SupportByVersion Excel 15,16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/jj227808.aspx </remarks>
        /// <param name="pivotItem">NetOffice.ExcelApi.PivotItem pivotItem</param>
        [CustomMethod]
        [SupportByVersion("Excel", 15, 16)]
        void DrillUp(NetOffice.ExcelApi.PivotItem pivotItem);

        /// <summary>
        /// SupportByVersion Excel 15,16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/jj227808.aspx </remarks>
        /// <param name="pivotItem">NetOffice.ExcelApi.PivotItem pivotItem</param>
        /// <param name="pivotLine">optional object pivotLine</param>
        [CustomMethod]
        [SupportByVersion("Excel", 15, 16)]
        void DrillUp(NetOffice.ExcelApi.PivotItem pivotItem, object pivotLine);

        /// <summary>
        /// SupportByVersion Excel 15,16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/jj230955.aspx </remarks>
        /// <param name="pivotItem">NetOffice.ExcelApi.PivotItem pivotItem</param>
        /// <param name="cubeField">NetOffice.ExcelApi.CubeField cubeField</param>
        /// <param name="pivotLine">optional object pivotLine</param>
        [SupportByVersion("Excel", 15, 16)]
        void DrillTo(NetOffice.ExcelApi.PivotItem pivotItem, NetOffice.ExcelApi.CubeField cubeField, object pivotLine);

        /// <summary>
        /// SupportByVersion Excel 15,16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/jj230955.aspx </remarks>
        /// <param name="pivotItem">NetOffice.ExcelApi.PivotItem pivotItem</param>
        /// <param name="cubeField">NetOffice.ExcelApi.CubeField cubeField</param>
        [CustomMethod]
        [SupportByVersion("Excel", 15, 16)]
        void DrillTo(NetOffice.ExcelApi.PivotItem pivotItem, NetOffice.ExcelApi.CubeField cubeField);

        /// <summary>
        /// SupportByVersion Excel 15,16
        /// </summary>
        /// <param name="arg1">object arg1</param>
        /// <param name="arg2">optional object arg2</param>
        /// <param name="arg3">optional object arg3</param>
        /// <param name="arg4">optional object arg4</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [SupportByVersion("Excel", 15, 16)]
        object Dummy2(object arg1, object arg2, object arg3, object arg4);

        /// <summary>
        /// SupportByVersion Excel 15,16
        /// </summary>
        /// <param name="arg1">object arg1</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [CustomMethod]
        [SupportByVersion("Excel", 15, 16)]
        object Dummy2(object arg1);

        /// <summary>
        /// SupportByVersion Excel 15,16
        /// </summary>
        /// <param name="arg1">object arg1</param>
        /// <param name="arg2">optional object arg2</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [CustomMethod]
        [SupportByVersion("Excel", 15, 16)]
        object Dummy2(object arg1, object arg2);

        /// <summary>
        /// SupportByVersion Excel 15,16
        /// </summary>
        /// <param name="arg1">object arg1</param>
        /// <param name="arg2">optional object arg2</param>
        /// <param name="arg3">optional object arg3</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [CustomMethod]
        [SupportByVersion("Excel", 15, 16)]
        object Dummy2(object arg1, object arg2, object arg3);

        #endregion
    }   
}
