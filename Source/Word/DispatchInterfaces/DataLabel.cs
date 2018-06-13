using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.WordApi
{
    /// <summary>
    /// DataLabel
    /// </summary>
    [SyntaxBypass]
    public interface DataLabel_ : ICOMObject
    {
        #region Properties

        /// <summary>
        /// SupportByVersion Word 14, 15, 16
        /// Get
        /// </summary>
        /// <param name="start">optional object start</param>
        /// <param name="length">optional object length</param>
        /// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff835187.aspx
        [SupportByVersion("Word", 14, 15, 16)]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        NetOffice.WordApi.ChartCharacters get_Characters(object start, object length);

        /// <summary>
        /// SupportByVersion Word 14, 15, 16
        /// Alias for get_Characters
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff835187.aspx </remarks>
        /// <param name="start">optional object start</param>
        /// <param name="length">optional object length</param>
        [SupportByVersion("Word", 14, 15, 16), Redirect("get_Characters")]
        NetOffice.WordApi.ChartCharacters Characters(object start, object length);

        /// <summary>
        /// SupportByVersion Word 14, 15, 16
        /// Get
        /// </summary>
        /// <param name="start">optional object start</param>
        /// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff835187.aspx
        [SupportByVersion("Word", 14, 15, 16)]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        NetOffice.WordApi.ChartCharacters get_Characters(object start);

        /// <summary>
        /// SupportByVersion Word 14, 15, 16
        /// Alias for get_Characters
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff835187.aspx </remarks>
        /// <param name="start">optional object start</param>
        [SupportByVersion("Word", 14, 15, 16), Redirect("get_Characters")]
        NetOffice.WordApi.ChartCharacters Characters(object start);

        #endregion
    }
    
    /// <summary>
    /// DispatchInterface DataLabel 
    /// SupportByVersion Word, 14,15,16
    /// </summary>
    /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff837261.aspx </remarks>
    [SupportByVersion("Word", 14, 15, 16)]
    [EntityType(EntityType.IsDispatchInterface)]
	[TypeId("1FD94DF1-3569-4465-94FF-E8B22D28EEB0")]
    public interface DataLabel : DataLabel_
    {
        #region Properties

        /// <summary>
        /// SupportByVersion Word 14, 15, 16
        /// Get
        /// Unknown COM Proxy
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff192189.aspx </remarks>
        [SupportByVersion("Word", 14, 15, 16), ProxyResult]
        object Parent { get; }

        /// <summary>
        /// SupportByVersion Word 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff836898.aspx </remarks>
        [SupportByVersion("Word", 14, 15, 16)]
        string Name { get; }

        /// <summary>
        /// SupportByVersion Word 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Word", 14, 15, 16)]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        NetOffice.WordApi.ChartBorder Border { get; }

        /// <summary>
        /// SupportByVersion Word 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Word", 14, 15, 16)]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        NetOffice.WordApi.Interior Interior { get; }

        /// <summary>
        /// SupportByVersion Word 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Word", 14, 15, 16)]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        NetOffice.WordApi.ChartFillFormat Fill { get; }

        /// <summary>
        /// SupportByVersion Word 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838756.aspx </remarks>
        [SupportByVersion("Word", 14, 15, 16)]
        string Caption { get; set; }

        /// <summary>
        /// SupportByVersion Word 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff835187.aspx </remarks>
        [SupportByVersion("Word", 14, 15, 16)]
        new NetOffice.WordApi.ChartCharacters Characters { get; }

        /// <summary>
        /// SupportByVersion Word 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Word", 14, 15, 16)]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        NetOffice.WordApi.ChartFont Font { get; }

        /// <summary>
        /// SupportByVersion Word 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff196941.aspx </remarks>
        [SupportByVersion("Word", 14, 15, 16)]
        object HorizontalAlignment { get; set; }

        /// <summary>
        /// SupportByVersion Word 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff192822.aspx </remarks>
        [SupportByVersion("Word", 14, 15, 16)]
        Double Left { get; set; }

        /// <summary>
        /// SupportByVersion Word 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff195292.aspx </remarks>
        [SupportByVersion("Word", 14, 15, 16)]
        object Orientation { get; set; }

        /// <summary>
        /// SupportByVersion Word 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff195979.aspx </remarks>
        [SupportByVersion("Word", 14, 15, 16)]
        bool Shadow { get; set; }

        /// <summary>
        /// SupportByVersion Word 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838149.aspx </remarks>
        [SupportByVersion("Word", 14, 15, 16)]
        string Text { get; set; }

        /// <summary>
        /// SupportByVersion Word 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194237.aspx </remarks>
        [SupportByVersion("Word", 14, 15, 16)]
        Double Top { get; set; }

        /// <summary>
        /// SupportByVersion Word 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff192774.aspx </remarks>
        [SupportByVersion("Word", 14, 15, 16)]
        object VerticalAlignment { get; set; }

        /// <summary>
        /// SupportByVersion Word 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff197902.aspx </remarks>
        [SupportByVersion("Word", 14, 15, 16)]
        Int32 ReadingOrder { get; set; }

        /// <summary>
        /// SupportByVersion Word 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Word", 14, 15, 16)]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        object AutoScaleFont { get; set; }

        /// <summary>
        /// SupportByVersion Word 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff839311.aspx </remarks>
        [SupportByVersion("Word", 14, 15, 16)]
        bool AutoText { get; set; }

        /// <summary>
        /// SupportByVersion Word 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff196392.aspx </remarks>
        [SupportByVersion("Word", 14, 15, 16)]
        string NumberFormat { get; set; }

        /// <summary>
        /// SupportByVersion Word 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff840631.aspx </remarks>
        [SupportByVersion("Word", 14, 15, 16)]
        bool NumberFormatLinked { get; set; }

        /// <summary>
        /// SupportByVersion Word 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff839142.aspx </remarks>
        [SupportByVersion("Word", 14, 15, 16)]
        object NumberFormatLocal { get; set; }

        /// <summary>
        /// SupportByVersion Word 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff837256.aspx </remarks>
        [SupportByVersion("Word", 14, 15, 16)]
        bool ShowLegendKey { get; set; }

        /// <summary>
        /// SupportByVersion Word 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Word", 14, 15, 16)]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        object Type { get; set; }

        /// <summary>
        /// SupportByVersion Word 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff197834.aspx </remarks>
        [SupportByVersion("Word", 14, 15, 16)]
        NetOffice.WordApi.Enums.XlDataLabelPosition Position { get; set; }

        /// <summary>
        /// SupportByVersion Word 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff197981.aspx </remarks>
        [SupportByVersion("Word", 14, 15, 16)]
        bool ShowSeriesName { get; set; }

        /// <summary>
        /// SupportByVersion Word 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff835520.aspx </remarks>
        [SupportByVersion("Word", 14, 15, 16)]
        bool ShowCategoryName { get; set; }

        /// <summary>
        /// SupportByVersion Word 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff192012.aspx </remarks>
        [SupportByVersion("Word", 14, 15, 16)]
        bool ShowValue { get; set; }

        /// <summary>
        /// SupportByVersion Word 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194599.aspx </remarks>
        [SupportByVersion("Word", 14, 15, 16)]
        bool ShowPercentage { get; set; }

        /// <summary>
        /// SupportByVersion Word 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff840450.aspx </remarks>
        [SupportByVersion("Word", 14, 15, 16)]
        bool ShowBubbleSize { get; set; }

        /// <summary>
        /// SupportByVersion Word 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff195601.aspx </remarks>
        [SupportByVersion("Word", 14, 15, 16)]
        object Separator { get; set; }

        /// <summary>
        /// SupportByVersion Word 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff840494.aspx </remarks>
        [SupportByVersion("Word", 14, 15, 16)]
        NetOffice.WordApi.ChartFormat Format { get; }

        /// <summary>
        /// SupportByVersion Word 14, 15, 16
        /// Get
        /// Unknown COM Proxy
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194158.aspx </remarks>
        [SupportByVersion("Word", 14, 15, 16), ProxyResult]
        object Application { get; }

        /// <summary>
        /// SupportByVersion Word 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194168.aspx </remarks>
        [SupportByVersion("Word", 14, 15, 16)]
        Int32 Creator { get; }

        /// <summary>
        /// SupportByVersion Word 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff198208.aspx </remarks>
        [SupportByVersion("Word", 14, 15, 16)]
        Double Height { get; set; }

        /// <summary>
        /// SupportByVersion Word 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff823008.aspx </remarks>
        [SupportByVersion("Word", 14, 15, 16)]
        Double Width { get; set; }

        /// <summary>
        /// SupportByVersion Word 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff835732.aspx </remarks>
        [SupportByVersion("Word", 14, 15, 16)]
        string Formula { get; set; }

        /// <summary>
        /// SupportByVersion Word 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff836621.aspx </remarks>
        [SupportByVersion("Word", 14, 15, 16)]
        string FormulaR1C1 { get; set; }

        /// <summary>
        /// SupportByVersion Word 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194198.aspx </remarks>
        [SupportByVersion("Word", 14, 15, 16)]
        string FormulaLocal { get; set; }

        /// <summary>
        /// SupportByVersion Word 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff840033.aspx </remarks>
        [SupportByVersion("Word", 14, 15, 16)]
        string FormulaR1C1Local { get; set; }

        /// <summary>
        /// SupportByVersion Word 15,16
        /// Get
        /// </summary>
        [SupportByVersion("Word", 15, 16)]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        Double _Height { get; }

        /// <summary>
        /// SupportByVersion Word 15,16
        /// Get
        /// </summary>
        [SupportByVersion("Word", 15, 16)]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        Double _Width { get; }

        /// <summary>
        /// SupportByVersion Word 15,16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/jj231842.aspx </remarks>
        [SupportByVersion("Word", 15, 16)]
        bool ShowRange { get; set; }

        #endregion

        #region Methods

        /// <summary>
        /// SupportByVersion Word 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff192051.aspx </remarks>
        [SupportByVersion("Word", 14, 15, 16)]
        object Select();

        /// <summary>
        /// SupportByVersion Word 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff840133.aspx </remarks>
        [SupportByVersion("Word", 14, 15, 16)]
        object Delete();

        #endregion
    }
}
