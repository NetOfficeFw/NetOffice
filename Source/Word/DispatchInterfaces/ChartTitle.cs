using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.WordApi
{
    /// <summary>
    /// ChartTitle
    /// </summary>
    [SyntaxBypass]
    public interface ChartTitle_ : ICOMObject
    {
        #region Properties

        /// <summary>
        /// SupportByVersion Word 14, 15, 16
        /// Get
        /// </summary>
        /// <param name="start">optional object start</param>
        /// <param name="length">optional object length</param>
        /// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff192392.aspx
        [SupportByVersion("Word", 14, 15, 16)]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        NetOffice.WordApi.ChartCharacters get_Characters(object start, object length);

        /// <summary>
        /// SupportByVersion Word 14, 15, 16
        /// Alias for get_Characters
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff192392.aspx </remarks>
        /// <param name="start">optional object start</param>
        /// <param name="length">optional object length</param>
        [SupportByVersion("Word", 14, 15, 16), Redirect("get_Characters")]
        NetOffice.WordApi.ChartCharacters Characters(object start, object length);

        /// <summary>
        /// SupportByVersion Word 14, 15, 16
        /// Get
        /// </summary>
        /// <param name="start">optional object start</param>
        /// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff192392.aspx
        [SupportByVersion("Word", 14, 15, 16)]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        NetOffice.WordApi.ChartCharacters get_Characters(object start);

        /// <summary>
        /// SupportByVersion Word 14, 15, 16
        /// Alias for get_Characters
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff192392.aspx </remarks>
        /// <param name="start">optional object start</param>
        [SupportByVersion("Word", 14, 15, 16), Redirect("get_Characters")]
        NetOffice.WordApi.ChartCharacters Characters(object start);

        #endregion
    }

    /// <summary>
    /// DispatchInterface ChartTitle 
    /// SupportByVersion Word, 14,15,16
    /// </summary>
    /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff840895.aspx </remarks>
    [SupportByVersion("Word", 14, 15, 16)]
    [EntityType(EntityType.IsDispatchInterface)]
	[TypeId("C1AD33E4-F088-40A9-9D2F-D94017D115C4")]
    public interface ChartTitle : ChartTitle_
    {
        #region Properties

        /// <summary>
        /// SupportByVersion Word 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff836026.aspx </remarks>
        [SupportByVersion("Word", 14, 15, 16)]
        string Caption { get; set; }

        /// <summary>
        /// SupportByVersion Word 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff192392.aspx </remarks>
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
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff196499.aspx </remarks>
        [SupportByVersion("Word", 14, 15, 16)]
        object HorizontalAlignment { get; set; }

        /// <summary>
        /// SupportByVersion Word 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff836107.aspx </remarks>
        [SupportByVersion("Word", 14, 15, 16)]
        Double Left { get; set; }

        /// <summary>
        /// SupportByVersion Word 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff822338.aspx </remarks>
        [SupportByVersion("Word", 14, 15, 16)]
        object Orientation { get; set; }

        /// <summary>
        /// SupportByVersion Word 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff192154.aspx </remarks>
        [SupportByVersion("Word", 14, 15, 16)]
        bool Shadow { get; set; }

        /// <summary>
        /// SupportByVersion Word 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff195363.aspx </remarks>
        [SupportByVersion("Word", 14, 15, 16)]
        string Text { get; set; }

        /// <summary>
        /// SupportByVersion Word 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff192759.aspx </remarks>
        [SupportByVersion("Word", 14, 15, 16)]
        Double Top { get; set; }

        /// <summary>
        /// SupportByVersion Word 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff845653.aspx </remarks>
        [SupportByVersion("Word", 14, 15, 16)]
        object VerticalAlignment { get; set; }

        /// <summary>
        /// SupportByVersion Word 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff836360.aspx </remarks>
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
        /// Get
        /// </summary>
        [SupportByVersion("Word", 14, 15, 16)]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        NetOffice.WordApi.ChartBorder Border { get; }

        /// <summary>
        /// SupportByVersion Word 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff837007.aspx </remarks>
        [SupportByVersion("Word", 14, 15, 16)]
        string Name { get; }

        /// <summary>
        /// SupportByVersion Word 14, 15, 16
        /// Get
        /// Unknown COM Proxy
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff197428.aspx </remarks>
        [SupportByVersion("Word", 14, 15, 16), ProxyResult]
        object Parent { get; }

        /// <summary>
        /// SupportByVersion Word 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff196531.aspx </remarks>
        [SupportByVersion("Word", 14, 15, 16)]
        bool IncludeInLayout { get; set; }

        /// <summary>
        /// SupportByVersion Word 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838760.aspx </remarks>
        [SupportByVersion("Word", 14, 15, 16)]
        NetOffice.WordApi.Enums.XlChartElementPosition Position { get; set; }

        /// <summary>
        /// SupportByVersion Word 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff197835.aspx </remarks>
        [SupportByVersion("Word", 14, 15, 16)]
        NetOffice.WordApi.ChartFormat Format { get; }

        /// <summary>
        /// SupportByVersion Word 14, 15, 16
        /// Get
        /// Unknown COM Proxy
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff836393.aspx </remarks>
        [SupportByVersion("Word", 14, 15, 16), ProxyResult]
        object Application { get; }

        /// <summary>
        /// SupportByVersion Word 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff841041.aspx </remarks>
        [SupportByVersion("Word", 14, 15, 16)]
        Int32 Creator { get; }

        /// <summary>
        /// SupportByVersion Word 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff197011.aspx </remarks>
        [SupportByVersion("Word", 14, 15, 16)]
        Double Height { get; }

        /// <summary>
        /// SupportByVersion Word 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff840079.aspx </remarks>
        [SupportByVersion("Word", 14, 15, 16)]
        Double Width { get; }

        /// <summary>
        /// SupportByVersion Word 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff839315.aspx </remarks>
        [SupportByVersion("Word", 14, 15, 16)]
        string Formula { get; set; }

        /// <summary>
        /// SupportByVersion Word 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff844811.aspx </remarks>
        [SupportByVersion("Word", 14, 15, 16)]
        string FormulaR1C1 { get; set; }

        /// <summary>
        /// SupportByVersion Word 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff845836.aspx </remarks>
        [SupportByVersion("Word", 14, 15, 16)]
        string FormulaLocal { get; set; }

        /// <summary>
        /// SupportByVersion Word 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838931.aspx </remarks>
        [SupportByVersion("Word", 14, 15, 16)]
        string FormulaR1C1Local { get; set; }

        #endregion

        #region Methods

        /// <summary>
        /// SupportByVersion Word 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff839489.aspx </remarks>
        [SupportByVersion("Word", 14, 15, 16)]
        object Delete();

        /// <summary>
        /// SupportByVersion Word 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff840563.aspx </remarks>
        [SupportByVersion("Word", 14, 15, 16)]
        object Select();

        #endregion
    }
}
