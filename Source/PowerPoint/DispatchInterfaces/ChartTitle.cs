using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.PowerPointApi
{
    /// <summary>
    /// ChartTitle
    /// </summary>
    [SyntaxBypass]
    public interface ChartTitle_ : ICOMObject
    {
        #region Properties

        /// <summary>
        /// SupportByVersion PowerPoint 14, 15, 16
        /// Get
        /// </summary>
        /// <param name="start">optional object start</param>
        /// <param name="length">optional object length</param>
        /// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff744571.aspx
        [SupportByVersion("PowerPoint", 14, 15, 16)]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        NetOffice.PowerPointApi.ChartCharacters get_Characters(object start, object length);

        /// <summary>
        /// SupportByVersion PowerPoint 14, 15, 16
        /// Alias for get_Characters
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff744571.aspx </remarks>
        /// <param name="start">optional object start</param>
        /// <param name="length">optional object length</param>
        [SupportByVersion("PowerPoint", 14, 15, 16), Redirect("get_Characters")]
        NetOffice.PowerPointApi.ChartCharacters Characters(object start, object length);

        /// <summary>
        /// SupportByVersion PowerPoint 14, 15, 16
        /// Get
        /// </summary>
        /// <param name="start">optional object start</param>
        /// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff744571.aspx
        [SupportByVersion("PowerPoint", 14, 15, 16)]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        NetOffice.PowerPointApi.ChartCharacters get_Characters(object start);

        /// <summary>
        /// SupportByVersion PowerPoint 14, 15, 16
        /// Alias for get_Characters
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff744571.aspx </remarks>
        /// <param name="start">optional object start</param>
        [SupportByVersion("PowerPoint", 14, 15, 16), Redirect("get_Characters")]
        NetOffice.PowerPointApi.ChartCharacters Characters(object start);

        #endregion
    }

    /// <summary>
    /// DispatchInterface ChartTitle
    /// SupportByVersion PowerPoint, 14,15,16
    /// </summary>
    /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff744207.aspx </remarks>
    [SupportByVersion("PowerPoint", 14, 15, 16)]
    [EntityType(EntityType.IsDispatchInterface)]
	[TypeId("92D41A5F-F07E-4CA4-AF6F-BEF486AA4E6F")]
    public interface ChartTitle : ChartTitle_
    {
        #region Properties

        /// <summary>
        /// SupportByVersion PowerPoint 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff745692.aspx </remarks>
        [SupportByVersion("PowerPoint", 14, 15, 16)]
        string Caption { get; set; }

        /// <summary>
        /// SupportByVersion PowerPoint 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff744571.aspx </remarks>
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
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff745167.aspx </remarks>
        [SupportByVersion("PowerPoint", 14, 15, 16)]
        object HorizontalAlignment { get; set; }

        /// <summary>
        /// SupportByVersion PowerPoint 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff745924.aspx </remarks>
        [SupportByVersion("PowerPoint", 14, 15, 16)]
        Double Left { get; set; }

        /// <summary>
        /// SupportByVersion PowerPoint 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff744839.aspx </remarks>
        [SupportByVersion("PowerPoint", 14, 15, 16)]
        object Orientation { get; set; }

        /// <summary>
        /// SupportByVersion PowerPoint 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff743978.aspx </remarks>
        [SupportByVersion("PowerPoint", 14, 15, 16)]
        bool Shadow { get; set; }

        /// <summary>
        /// SupportByVersion PowerPoint 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff743850.aspx </remarks>
        [SupportByVersion("PowerPoint", 14, 15, 16)]
        string Text { get; set; }

        /// <summary>
        /// SupportByVersion PowerPoint 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff746036.aspx </remarks>
        [SupportByVersion("PowerPoint", 14, 15, 16)]
        Double Top { get; set; }

        /// <summary>
        /// SupportByVersion PowerPoint 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff744984.aspx </remarks>
        [SupportByVersion("PowerPoint", 14, 15, 16)]
        object VerticalAlignment { get; set; }

        /// <summary>
        /// SupportByVersion PowerPoint 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("PowerPoint", 14, 15, 16)]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        object AutoScaleFont { get; set; }

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
        /// Get
        /// </summary>
        [SupportByVersion("PowerPoint", 14, 15, 16)]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        NetOffice.PowerPointApi.ChartBorder Border { get; }

        /// <summary>
        /// SupportByVersion PowerPoint 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff746337.aspx </remarks>
        [SupportByVersion("PowerPoint", 14, 15, 16)]
        string Name { get; }

        /// <summary>
        /// SupportByVersion PowerPoint 14, 15, 16
        /// Get
        /// Unknown COM Proxy
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff746701.aspx </remarks>
        [SupportByVersion("PowerPoint", 14, 15, 16), ProxyResult]
        object Parent { get; }

        /// <summary>
        /// SupportByVersion PowerPoint 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff746366.aspx </remarks>
        [SupportByVersion("PowerPoint", 14, 15, 16)]
        bool IncludeInLayout { get; set; }

        /// <summary>
        /// SupportByVersion PowerPoint 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff745966.aspx </remarks>
        [SupportByVersion("PowerPoint", 14, 15, 16)]
        NetOffice.PowerPointApi.Enums.XlChartElementPosition Position { get; set; }

        /// <summary>
        /// SupportByVersion PowerPoint 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff746439.aspx </remarks>
        [SupportByVersion("PowerPoint", 14, 15, 16)]
        NetOffice.PowerPointApi.ChartFormat Format { get; }

        /// <summary>
        /// SupportByVersion PowerPoint 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff745520.aspx </remarks>
        [SupportByVersion("PowerPoint", 14, 15, 16)]
        Int32 Creator { get; }

        /// <summary>
        /// SupportByVersion PowerPoint 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff745769.aspx </remarks>
        [SupportByVersion("PowerPoint", 14, 15, 16)]
        NetOffice.PowerPointApi.Application Application { get; }

        /// <summary>
        /// SupportByVersion PowerPoint 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff745321.aspx </remarks>
        [SupportByVersion("PowerPoint", 14, 15, 16)]
        Int32 ReadingOrder { get; set; }

        /// <summary>
        /// SupportByVersion PowerPoint 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff746385.aspx </remarks>
        [SupportByVersion("PowerPoint", 14, 15, 16)]
        Double Height { get; }

        /// <summary>
        /// SupportByVersion PowerPoint 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff746776.aspx </remarks>
        [SupportByVersion("PowerPoint", 14, 15, 16)]
        Double Width { get; }

        /// <summary>
        /// SupportByVersion PowerPoint 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff744581.aspx </remarks>
        [SupportByVersion("PowerPoint", 14, 15, 16)]
        string Formula { get; set; }

        /// <summary>
        /// SupportByVersion PowerPoint 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff746523.aspx </remarks>
        [SupportByVersion("PowerPoint", 14, 15, 16)]
        string FormulaR1C1 { get; set; }

        /// <summary>
        /// SupportByVersion PowerPoint 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff745539.aspx </remarks>
        [SupportByVersion("PowerPoint", 14, 15, 16)]
        string FormulaLocal { get; set; }

        /// <summary>
        /// SupportByVersion PowerPoint 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff745254.aspx </remarks>
        [SupportByVersion("PowerPoint", 14, 15, 16)]
        string FormulaR1C1Local { get; set; }

        #endregion

        #region Methods

        /// <summary>
        /// SupportByVersion PowerPoint 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff745457.aspx </remarks>
        [SupportByVersion("PowerPoint", 14, 15, 16)]
        object Delete();

        /// <summary>
        /// SupportByVersion PowerPoint 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff745210.aspx </remarks>
        [SupportByVersion("PowerPoint", 14, 15, 16)]
        object Select();

        #endregion
    }
}
