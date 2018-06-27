using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.PowerPointApi
{
    /// <summary>
    /// AxisTitle
    /// </summary>
    [SyntaxBypass]
    public interface AxisTitle_ : ICOMObject
    {
        #region Properties

        /// <summary>
        /// SupportByVersion PowerPoint 14, 15, 16
        /// Get
        /// </summary>
        /// <param name="start">optional object start</param>
        /// <param name="length">optional object length</param>
        /// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff745546.aspx
        [SupportByVersion("PowerPoint", 14, 15, 16)]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        NetOffice.PowerPointApi.ChartCharacters get_Characters(object start, object length);

        /// <summary>
        /// SupportByVersion PowerPoint 14, 15, 16
        /// Alias for get_Characters
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff745546.aspx </remarks>
        /// <param name="start">optional object start</param>
        /// <param name="length">optional object length</param>
        [SupportByVersion("PowerPoint", 14, 15, 16), Redirect("get_Characters")]
        NetOffice.PowerPointApi.ChartCharacters Characters(object start, object length);

        /// <summary>
        /// SupportByVersion PowerPoint 14, 15, 16
        /// Get
        /// </summary>
        /// <param name="start">optional object start</param>
        /// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff745546.aspx
        [SupportByVersion("PowerPoint", 14, 15, 16)]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        NetOffice.PowerPointApi.ChartCharacters get_Characters(object start);

        /// <summary>
        /// SupportByVersion PowerPoint 14, 15, 16
        /// Alias for get_Characters
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff745546.aspx </remarks>
        /// <param name="start">optional object start</param>
        [SupportByVersion("PowerPoint", 14, 15, 16), Redirect("get_Characters")]
        NetOffice.PowerPointApi.ChartCharacters Characters(object start);

        #endregion
    }

    /// <summary>
    /// DispatchInterface AxisTitle
    /// SupportByVersion PowerPoint, 14,15,16
    /// </summary>
    /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff745597.aspx </remarks>
    [SupportByVersion("PowerPoint", 14, 15, 16)]
    [EntityType(EntityType.IsDispatchInterface)]
	[TypeId("92D41A54-F07E-4CA4-AF6F-BEF486AA4E6F")]
    public interface AxisTitle : AxisTitle_
    {
        #region Properties

        /// <summary>
        /// SupportByVersion PowerPoint 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff746628.aspx </remarks>
        [SupportByVersion("PowerPoint", 14, 15, 16)]
        string Caption { get; set; }

        /// <summary>
        /// SupportByVersion PowerPoint 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff745546.aspx </remarks>
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
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff746675.aspx </remarks>
        [SupportByVersion("PowerPoint", 14, 15, 16)]
        object HorizontalAlignment { get; set; }

        /// <summary>
        /// SupportByVersion PowerPoint 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff744140.aspx </remarks>
        [SupportByVersion("PowerPoint", 14, 15, 16)]
        Double Left { get; set; }

        /// <summary>
        /// SupportByVersion PowerPoint 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff746033.aspx </remarks>
        [SupportByVersion("PowerPoint", 14, 15, 16)]
        object Orientation { get; set; }

        /// <summary>
        /// SupportByVersion PowerPoint 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff744568.aspx </remarks>
        [SupportByVersion("PowerPoint", 14, 15, 16)]
        bool Shadow { get; set; }

        /// <summary>
        /// SupportByVersion PowerPoint 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff746196.aspx </remarks>
        [SupportByVersion("PowerPoint", 14, 15, 16)]
        string Text { get; set; }

        /// <summary>
        /// SupportByVersion PowerPoint 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff744545.aspx </remarks>
        [SupportByVersion("PowerPoint", 14, 15, 16)]
        Double Top { get; set; }

        /// <summary>
        /// SupportByVersion PowerPoint 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff743882.aspx </remarks>
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
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff746692.aspx </remarks>
        [SupportByVersion("PowerPoint", 14, 15, 16)]
        string Name { get; }

        /// <summary>
        /// SupportByVersion PowerPoint 14, 15, 16
        /// Get
        /// Unknown COM Proxy
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff745911.aspx </remarks>
        [SupportByVersion("PowerPoint", 14, 15, 16), ProxyResult]
        object Parent { get; }

        /// <summary>
        /// SupportByVersion PowerPoint 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff743946.aspx </remarks>
        [SupportByVersion("PowerPoint", 14, 15, 16)]
        bool IncludeInLayout { get; set; }

        /// <summary>
        /// SupportByVersion PowerPoint 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff745347.aspx </remarks>
        [SupportByVersion("PowerPoint", 14, 15, 16)]
        NetOffice.PowerPointApi.Enums.XlChartElementPosition Position { get; set; }

        /// <summary>
        /// SupportByVersion PowerPoint 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff745870.aspx </remarks>
        [SupportByVersion("PowerPoint", 14, 15, 16)]
        NetOffice.PowerPointApi.ChartFormat Format { get; }

        /// <summary>
        /// SupportByVersion PowerPoint 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff746782.aspx </remarks>
        [SupportByVersion("PowerPoint", 14, 15, 16)]
        Int32 Creator { get; }

        /// <summary>
        /// SupportByVersion PowerPoint 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff746190.aspx </remarks>
        [SupportByVersion("PowerPoint", 14, 15, 16)]
        NetOffice.PowerPointApi.Application Application { get; }

        /// <summary>
        /// SupportByVersion PowerPoint 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff744796.aspx </remarks>
        [SupportByVersion("PowerPoint", 14, 15, 16)]
        Int32 ReadingOrder { get; set; }

        /// <summary>
        /// SupportByVersion PowerPoint 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff746563.aspx </remarks>
        [SupportByVersion("PowerPoint", 14, 15, 16)]
        Double Height { get; }

        /// <summary>
        /// SupportByVersion PowerPoint 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff745147.aspx </remarks>
        [SupportByVersion("PowerPoint", 14, 15, 16)]
        Double Width { get; }

        /// <summary>
        /// SupportByVersion PowerPoint 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff745933.aspx </remarks>
        [SupportByVersion("PowerPoint", 14, 15, 16)]
        string Formula { get; set; }

        /// <summary>
        /// SupportByVersion PowerPoint 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff744316.aspx </remarks>
        [SupportByVersion("PowerPoint", 14, 15, 16)]
        string FormulaR1C1 { get; set; }

        /// <summary>
        /// SupportByVersion PowerPoint 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff745627.aspx </remarks>
        [SupportByVersion("PowerPoint", 14, 15, 16)]
        string FormulaLocal { get; set; }

        /// <summary>
        /// SupportByVersion PowerPoint 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff744715.aspx </remarks>
        [SupportByVersion("PowerPoint", 14, 15, 16)]
        string FormulaR1C1Local { get; set; }

        #endregion

        #region Methods

        /// <summary>
        /// SupportByVersion PowerPoint 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff744056.aspx </remarks>
        [SupportByVersion("PowerPoint", 14, 15, 16)]
        object Delete();

        /// <summary>
        /// SupportByVersion PowerPoint 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff744775.aspx </remarks>
        [SupportByVersion("PowerPoint", 14, 15, 16)]
        object Select();

        #endregion
    }
}
