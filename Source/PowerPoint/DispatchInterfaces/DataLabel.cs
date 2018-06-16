using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.PowerPointApi
{
    /// <summary>
    /// DataLabel
    /// </summary>
    [SyntaxBypass]
    public interface DataLabel_ : ICOMObject
    {
        #region Properties

        /// <summary>
        /// SupportByVersion PowerPoint 14, 15, 16
        /// Get
        /// </summary>
        /// <param name="start">optional object start</param>
        /// <param name="length">optional object length</param>
        /// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff743958.aspx
        [SupportByVersion("PowerPoint", 14, 15, 16)]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        NetOffice.PowerPointApi.ChartCharacters get_Characters(object start, object length);

        /// <summary>
        /// SupportByVersion PowerPoint 14, 15, 16
        /// Alias for get_Characters
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff743958.aspx </remarks>
        /// <param name="start">optional object start</param>
        /// <param name="length">optional object length</param>
        [SupportByVersion("PowerPoint", 14, 15, 16), Redirect("get_Characters")]
        NetOffice.PowerPointApi.ChartCharacters Characters(object start, object length);

        /// <summary>
        /// SupportByVersion PowerPoint 14, 15, 16
        /// Get
        /// </summary>
        /// <param name="start">optional object start</param>
        /// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff743958.aspx
        [SupportByVersion("PowerPoint", 14, 15, 16)]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        NetOffice.PowerPointApi.ChartCharacters get_Characters(object start);

        /// <summary>
        /// SupportByVersion PowerPoint 14, 15, 16
        /// Alias for get_Characters
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff743958.aspx </remarks>
        /// <param name="start">optional object start</param>
        [SupportByVersion("PowerPoint", 14, 15, 16), Redirect("get_Characters")]
        NetOffice.PowerPointApi.ChartCharacters Characters(object start);

        #endregion
    }

    /// <summary>
    /// DispatchInterface DataLabel 
    /// SupportByVersion PowerPoint, 14,15,16
    /// </summary>
    /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff745813.aspx </remarks>
    [SupportByVersion("PowerPoint", 14, 15, 16)]
    [EntityType(EntityType.IsDispatchInterface)]
	[TypeId("92D41A61-F07E-4CA4-AF6F-BEF486AA4E6F")]
    public interface DataLabel : DataLabel_
    {
        #region Properties

        /// <summary>
        /// SupportByVersion PowerPoint 14, 15, 16
        /// Get
        /// Unknown COM Proxy
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff745516.aspx </remarks>
        [SupportByVersion("PowerPoint", 14, 15, 16), ProxyResult]
        object Parent { get; }

        /// <summary>
        /// SupportByVersion PowerPoint 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff744660.aspx </remarks>
        [SupportByVersion("PowerPoint", 14, 15, 16)]
        string Name { get; }

        /// <summary>
        /// SupportByVersion PowerPoint 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("PowerPoint", 14, 15, 16)]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        NetOffice.PowerPointApi.ChartBorder Border { get; }

        /// <summary>
        /// SupportByVersion PowerPoint 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("PowerPoint", 14, 15, 16)]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        NetOffice.PowerPointApi.Interior Interior { get; }

        /// <summary>
        /// SupportByVersion PowerPoint 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("PowerPoint", 14, 15, 16)]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        NetOffice.PowerPointApi.ChartFillFormat Fill { get; }

        /// <summary>
        /// SupportByVersion PowerPoint 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff746038.aspx </remarks>
        [SupportByVersion("PowerPoint", 14, 15, 16)]
        string Caption { get; set; }

        /// <summary>
        /// SupportByVersion PowerPoint 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff743958.aspx </remarks>
        [SupportByVersion("PowerPoint", 14, 15, 16)]
        new NetOffice.PowerPointApi.ChartCharacters Characters { get; }

        /// <summary>
        /// SupportByVersion PowerPoint 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("PowerPoint", 14, 15, 16)]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        NetOffice.PowerPointApi.ChartFont Font { get; }

        /// <summary>
        /// SupportByVersion PowerPoint 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff744126.aspx </remarks>
        [SupportByVersion("PowerPoint", 14, 15, 16)]
        object HorizontalAlignment { get; set; }

        /// <summary>
        /// SupportByVersion PowerPoint 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff745315.aspx </remarks>
        [SupportByVersion("PowerPoint", 14, 15, 16)]
        Double Left { get; set; }

        /// <summary>
        /// SupportByVersion PowerPoint 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff744363.aspx </remarks>
        [SupportByVersion("PowerPoint", 14, 15, 16)]
        object Orientation { get; set; }

        /// <summary>
        /// SupportByVersion PowerPoint 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff746270.aspx </remarks>
        [SupportByVersion("PowerPoint", 14, 15, 16)]
        bool Shadow { get; set; }

        /// <summary>
        /// SupportByVersion PowerPoint 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff746797.aspx </remarks>
        [SupportByVersion("PowerPoint", 14, 15, 16)]
        string Text { get; set; }

        /// <summary>
        /// SupportByVersion PowerPoint 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff746355.aspx </remarks>
        [SupportByVersion("PowerPoint", 14, 15, 16)]
        Double Top { get; set; }

        /// <summary>
        /// SupportByVersion PowerPoint 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff744338.aspx </remarks>
        [SupportByVersion("PowerPoint", 14, 15, 16)]
        object VerticalAlignment { get; set; }

        /// <summary>
        /// SupportByVersion PowerPoint 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff745855.aspx </remarks>
        [SupportByVersion("PowerPoint", 14, 15, 16)]
        Int32 ReadingOrder { get; set; }

        /// <summary>
        /// SupportByVersion PowerPoint 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("PowerPoint", 14, 15, 16)]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        object AutoScaleFont { get; set; }

        /// <summary>
        /// SupportByVersion PowerPoint 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff746751.aspx </remarks>
        [SupportByVersion("PowerPoint", 14, 15, 16)]
        bool AutoText { get; set; }

        /// <summary>
        /// SupportByVersion PowerPoint 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff745971.aspx </remarks>
        [SupportByVersion("PowerPoint", 14, 15, 16)]
        string NumberFormat { get; set; }

        /// <summary>
        /// SupportByVersion PowerPoint 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff745693.aspx </remarks>
        [SupportByVersion("PowerPoint", 14, 15, 16)]
        bool NumberFormatLinked { get; set; }

        /// <summary>
        /// SupportByVersion PowerPoint 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff744772.aspx </remarks>
        [SupportByVersion("PowerPoint", 14, 15, 16)]
        object NumberFormatLocal { get; set; }

        /// <summary>
        /// SupportByVersion PowerPoint 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff744166.aspx </remarks>
        [SupportByVersion("PowerPoint", 14, 15, 16)]
        bool ShowLegendKey { get; set; }

        /// <summary>
        /// SupportByVersion PowerPoint 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("PowerPoint", 14, 15, 16)]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        object Type { get; set; }

        /// <summary>
        /// SupportByVersion PowerPoint 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff744685.aspx </remarks>
        [SupportByVersion("PowerPoint", 14, 15, 16)]
        NetOffice.PowerPointApi.Enums.XlDataLabelPosition Position { get; set; }

        /// <summary>
        /// SupportByVersion PowerPoint 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff744996.aspx </remarks>
        [SupportByVersion("PowerPoint", 14, 15, 16)]
        bool ShowSeriesName { get; set; }

        /// <summary>
        /// SupportByVersion PowerPoint 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff745392.aspx </remarks>
        [SupportByVersion("PowerPoint", 14, 15, 16)]
        bool ShowCategoryName { get; set; }

        /// <summary>
        /// SupportByVersion PowerPoint 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff744353.aspx </remarks>
        [SupportByVersion("PowerPoint", 14, 15, 16)]
        bool ShowValue { get; set; }

        /// <summary>
        /// SupportByVersion PowerPoint 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff743836.aspx </remarks>
        [SupportByVersion("PowerPoint", 14, 15, 16)]
        bool ShowPercentage { get; set; }

        /// <summary>
        /// SupportByVersion PowerPoint 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff745857.aspx </remarks>
        [SupportByVersion("PowerPoint", 14, 15, 16)]
        bool ShowBubbleSize { get; set; }

        /// <summary>
        /// SupportByVersion PowerPoint 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff744090.aspx </remarks>
        [SupportByVersion("PowerPoint", 14, 15, 16)]
        object Separator { get; set; }

        /// <summary>
        /// SupportByVersion PowerPoint 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff746652.aspx </remarks>
        [SupportByVersion("PowerPoint", 14, 15, 16)]
        NetOffice.PowerPointApi.ChartFormat Format { get; }

        /// <summary>
        /// SupportByVersion PowerPoint 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff744299.aspx </remarks>
        [SupportByVersion("PowerPoint", 14, 15, 16)]
        Int32 Creator { get; }

        /// <summary>
        /// SupportByVersion PowerPoint 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff745389.aspx </remarks>
        [SupportByVersion("PowerPoint", 14, 15, 16)]
        NetOffice.PowerPointApi.Application Application { get; }

        /// <summary>
        /// SupportByVersion PowerPoint 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff745512.aspx </remarks>
        [SupportByVersion("PowerPoint", 14, 15, 16)]
        Double Height { get; set; }

        /// <summary>
        /// SupportByVersion PowerPoint 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff744304.aspx </remarks>
        [SupportByVersion("PowerPoint", 14, 15, 16)]
        Double Width { get; set; }

        /// <summary>
        /// SupportByVersion PowerPoint 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff746432.aspx </remarks>
        [SupportByVersion("PowerPoint", 14, 15, 16)]
        string Formula { get; set; }

        /// <summary>
        /// SupportByVersion PowerPoint 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff746504.aspx </remarks>
        [SupportByVersion("PowerPoint", 14, 15, 16)]
        string FormulaR1C1 { get; set; }

        /// <summary>
        /// SupportByVersion PowerPoint 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff746101.aspx </remarks>
        [SupportByVersion("PowerPoint", 14, 15, 16)]
        string FormulaLocal { get; set; }

        /// <summary>
        /// SupportByVersion PowerPoint 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff744766.aspx </remarks>
        [SupportByVersion("PowerPoint", 14, 15, 16)]
        string FormulaR1C1Local { get; set; }

        /// <summary>
        /// SupportByVersion PowerPoint 15,16
        /// Get
        /// </summary>
        [SupportByVersion("PowerPoint", 15, 16)]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        Double _Height { get; }

        /// <summary>
        /// SupportByVersion PowerPoint 15,16
        /// Get
        /// </summary>
        [SupportByVersion("PowerPoint", 15, 16)]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        Double _Width { get; }

        /// <summary>
        /// SupportByVersion PowerPoint 15,16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/jj228303.aspx </remarks>
        [SupportByVersion("PowerPoint", 15, 16)]
        bool ShowRange { get; set; }

        #endregion

        #region Methods

        /// <summary>
        /// SupportByVersion PowerPoint 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff744031.aspx </remarks>
        [SupportByVersion("PowerPoint", 14, 15, 16)]
        object Select();

        /// <summary>
        /// SupportByVersion PowerPoint 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff745601.aspx </remarks>
        [SupportByVersion("PowerPoint", 14, 15, 16)]
        object Delete();

        #endregion
    }  
}
