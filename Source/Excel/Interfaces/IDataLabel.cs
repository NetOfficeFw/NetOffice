using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.ExcelApi
{
    /// <summary>
    /// IDataLabel
    /// </summary>
    [SyntaxBypass]
    public interface IDataLabel_ : ICOMObject
    {
        #region Properties

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <param name="start">optional object start</param>
        /// <param name="length">optional object length</param>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        NetOffice.ExcelApi.Characters get_Characters(object start, object length);

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Alias for get_Characters
        /// </summary>
        /// <param name="start">optional object start</param>
        /// <param name="length">optional object length</param>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16), Redirect("get_Characters")]
        NetOffice.ExcelApi.Characters Characters(object start, object length);

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <param name="start">optional object start</param>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        NetOffice.ExcelApi.Characters get_Characters(object start);

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Alias for get_Characters
        /// </summary>
        /// <param name="start">optional object start</param>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16), Redirect("get_Characters")]
        NetOffice.ExcelApi.Characters Characters(object start);

        #endregion
    }

    /// <summary>
    /// Interface IDataLabel 
    /// SupportByVersion Excel, 9,10,11,12,14,15,16
    /// </summary>
    [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
    [EntityType(EntityType.IsInterface)]
    public interface IDataLabel : IDataLabel_
    {
        #region Properties

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        NetOffice.ExcelApi.Application Application { get; }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        NetOffice.ExcelApi.Enums.XlCreator Creator { get; }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// Unknown COM Proxy
        /// </summary>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16), ProxyResult]
        object Parent { get; }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        string Name { get; }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        NetOffice.ExcelApi.Border Border { get; }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        NetOffice.ExcelApi.Interior Interior { get; }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        NetOffice.ExcelApi.ChartFillFormat Fill { get; }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        string Caption { get; set; }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        NetOffice.ExcelApi.Characters Characters { get; }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        NetOffice.ExcelApi.Font Font { get; }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        object HorizontalAlignment { get; set; }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        Double Left { get; set; }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        object Orientation { get; set; }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        bool Shadow { get; set; }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        string Text { get; set; }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        Double Top { get; set; }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        object VerticalAlignment { get; set; }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        Int32 ReadingOrder { get; set; }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        object AutoScaleFont { get; set; }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        bool AutoText { get; set; }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        string NumberFormat { get; set; }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        bool NumberFormatLinked { get; set; }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        object NumberFormatLocal { get; set; }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        bool ShowLegendKey { get; set; }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        object Type { get; set; }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        NetOffice.ExcelApi.Enums.XlDataLabelPosition Position { get; set; }

        /// <summary>
        /// SupportByVersion Excel 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Excel", 10, 11, 12, 14, 15, 16)]
        bool ShowSeriesName { get; set; }

        /// <summary>
        /// SupportByVersion Excel 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Excel", 10, 11, 12, 14, 15, 16)]
        bool ShowCategoryName { get; set; }

        /// <summary>
        /// SupportByVersion Excel 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Excel", 10, 11, 12, 14, 15, 16)]
        bool ShowValue { get; set; }

        /// <summary>
        /// SupportByVersion Excel 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Excel", 10, 11, 12, 14, 15, 16)]
        bool ShowPercentage { get; set; }

        /// <summary>
        /// SupportByVersion Excel 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Excel", 10, 11, 12, 14, 15, 16)]
        bool ShowBubbleSize { get; set; }

        /// <summary>
        /// SupportByVersion Excel 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Excel", 10, 11, 12, 14, 15, 16)]
        object Separator { get; set; }

        /// <summary>
        /// SupportByVersion Excel 12, 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Excel", 12, 14, 15, 16)]
        NetOffice.ExcelApi.ChartFormat Format { get; }

        /// <summary>
        /// SupportByVersion Excel 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Excel", 14, 15, 16)]
        Double Height { get; set; }

        /// <summary>
        /// SupportByVersion Excel 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Excel", 14, 15, 16)]
        Double Width { get; set; }

        /// <summary>
        /// SupportByVersion Excel 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Excel", 14, 15, 16)]
        string Formula { get; set; }

        /// <summary>
        /// SupportByVersion Excel 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Excel", 14, 15, 16)]
        string FormulaR1C1 { get; set; }

        /// <summary>
        /// SupportByVersion Excel 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Excel", 14, 15, 16)]
        string FormulaLocal { get; set; }

        /// <summary>
        /// SupportByVersion Excel 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Excel", 14, 15, 16)]
        string FormulaR1C1Local { get; set; }

        /// <summary>
        /// SupportByVersion Excel 15,16
        /// Get
        /// </summary>
        [SupportByVersion("Excel", 15, 16)]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        Double _Height { get; }

        /// <summary>
        /// SupportByVersion Excel 15,16
        /// Get
        /// </summary>
        [SupportByVersion("Excel", 15, 16)]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        Double _Width { get; }

        /// <summary>
        /// SupportByVersion Excel 15,16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Excel", 15, 16)]
        bool ShowRange { get; set; }

        #endregion

        #region Methods

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        object Select();

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        object Delete();

        #endregion
    }
}
