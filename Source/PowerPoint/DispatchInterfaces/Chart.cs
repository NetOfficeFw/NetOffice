using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.PowerPointApi
{
    /// <summary>
    /// Chart
    /// </summary>
    [SyntaxBypass]
    public interface Chart_ : ICOMObject
    {
        #region Properties

        /// <summary>
        /// SupportByVersion PowerPoint 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <param name="index1">optional object index1</param>
        /// <param name="index2">optional object index2</param>
        /// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff746650.aspx
        [SupportByVersion("PowerPoint", 14, 15, 16)]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        object get_HasAxis(object index1, object index2);

        /// <summary>
        /// SupportByVersion PowerPoint 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <param name="index1">optional object index1</param>
        /// <param name="index2">optional object index2</param>
        /// <param name="value">optional object value</param>
        [SupportByVersion("PowerPoint", 14, 15, 16)]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        void set_HasAxis(object index1, object index2, object value);

        /// <summary>
        /// SupportByVersion PowerPoint 14, 15, 16
        /// Alias for get_HasAxis
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff746650.aspx </remarks>
        /// <param name="index1">optional object index1</param>
        /// <param name="index2">optional object index2</param>
        [SupportByVersion("PowerPoint", 14, 15, 16), Redirect("get_HasAxis")]
        object HasAxis(object index1, object index2);

        /// <summary>
        /// SupportByVersion PowerPoint 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <param name="index1">optional object index1</param>
        /// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff746650.aspx
        [SupportByVersion("PowerPoint", 14, 15, 16)]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        object get_HasAxis(object index1);

        /// <summary>
        /// SupportByVersion PowerPoint 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <param name="index1">optional object index1</param>
        /// <param name="value">optional object value</param>
        [SupportByVersion("PowerPoint", 14, 15, 16)]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        void set_HasAxis(object index1, object value);

        /// <summary>
        /// SupportByVersion PowerPoint 14, 15, 16
        /// Alias for get_HasAxis
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff746650.aspx </remarks>
        /// <param name="index1">optional object index1</param>
        [SupportByVersion("PowerPoint", 14, 15, 16), Redirect("get_HasAxis")]
        object HasAxis(object index1);

        #endregion
    }
   
    /// <summary>
    /// DispatchInterface Chart 
    /// SupportByVersion PowerPoint, 14,15,16
    /// </summary>
    /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff744663.aspx </remarks>
    [SupportByVersion("PowerPoint", 14, 15, 16)]
    [EntityType(EntityType.IsDispatchInterface)]
	[TypeId("92D41A55-F07E-4CA4-AF6F-BEF486AA4E6F")]
    public interface Chart : Chart_
    {
        #region Properties

        /// <summary>
        /// SupportByVersion PowerPoint 14, 15, 16
        /// Get
        /// Unknown COM Proxy
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff746116.aspx </remarks>
        [SupportByVersion("PowerPoint", 14, 15, 16), ProxyResult]
        object Parent { get; }

        /// <summary>
        /// SupportByVersion PowerPoint 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff744954.aspx </remarks>
        [SupportByVersion("PowerPoint", 14, 15, 16)]
        NetOffice.OfficeApi.Enums.XlChartType ChartType { get; set; }

        /// <summary>
        /// SupportByVersion PowerPoint 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff745140.aspx </remarks>
        [SupportByVersion("PowerPoint", 14, 15, 16)]
        bool HasDataTable { get; set; }

        /// <summary>
        /// SupportByVersion PowerPoint 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff744071.aspx </remarks>
        [SupportByVersion("PowerPoint", 14, 15, 16)]
        NetOffice.PowerPointApi.Enums.XlRowCol PlotBy { get; set; }

        /// <summary>
        /// SupportByVersion PowerPoint 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff746809.aspx </remarks>
        [SupportByVersion("PowerPoint", 14, 15, 16)]
        NetOffice.PowerPointApi.DataTable DataTable { get; }

        /// <summary>
        /// SupportByVersion PowerPoint 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff746790.aspx </remarks>
        [SupportByVersion("PowerPoint", 14, 15, 16)]
        NetOffice.PowerPointApi.Enums.XlBarShape BarShape { get; set; }

        /// <summary>
        /// SupportByVersion PowerPoint 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff744381.aspx </remarks>
        [SupportByVersion("PowerPoint", 14, 15, 16)]
        NetOffice.PowerPointApi.Walls SideWall { get; }

        /// <summary>
        /// SupportByVersion PowerPoint 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff744079.aspx </remarks>
        [SupportByVersion("PowerPoint", 14, 15, 16)]
        NetOffice.PowerPointApi.Walls BackWall { get; }

        /// <summary>
        /// SupportByVersion PowerPoint 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff743954.aspx </remarks>
        [SupportByVersion("PowerPoint", 14, 15, 16)]
        object ChartStyle { get; set; }

        /// <summary>
        /// SupportByVersion PowerPoint 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("PowerPoint", 14, 15, 16)]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        bool HasPivotFields { get; set; }

        /// <summary>
        /// SupportByVersion PowerPoint 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff745647.aspx </remarks>
        [SupportByVersion("PowerPoint", 14, 15, 16)]
        bool ShowDataLabelsOverMaximum { get; set; }

        /// <summary>
        /// SupportByVersion PowerPoint 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff744089.aspx </remarks>
        [SupportByVersion("PowerPoint", 14, 15, 16)]
        NetOffice.PowerPointApi.ChartData ChartData { get; }

        /// <summary>
        /// SupportByVersion PowerPoint 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff746059.aspx </remarks>
        [SupportByVersion("PowerPoint", 14, 15, 16)]
        NetOffice.PowerPointApi.Shapes Shapes { get; }

        /// <summary>
        /// SupportByVersion PowerPoint 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff746336.aspx </remarks>
        [SupportByVersion("PowerPoint", 14, 15, 16)]
        Int32 Creator { get; }

        /// <summary>
        /// SupportByVersion PowerPoint 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("PowerPoint", 14, 15, 16)]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        NetOffice.PowerPointApi.ChartGroup Area3DGroup { get; }

        /// <summary>
        /// SupportByVersion PowerPoint 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("PowerPoint", 14, 15, 16)]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        NetOffice.PowerPointApi.ChartGroup Bar3DGroup { get; }

        /// <summary>
        /// SupportByVersion PowerPoint 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("PowerPoint", 14, 15, 16)]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        NetOffice.PowerPointApi.ChartGroup Column3DGroup { get; }

        /// <summary>
        /// SupportByVersion PowerPoint 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("PowerPoint", 14, 15, 16)]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        NetOffice.PowerPointApi.ChartGroup Line3DGroup { get; }

        /// <summary>
        /// SupportByVersion PowerPoint 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("PowerPoint", 14, 15, 16)]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        NetOffice.PowerPointApi.ChartGroup Pie3DGroup { get; }

        /// <summary>
        /// SupportByVersion PowerPoint 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("PowerPoint", 14, 15, 16)]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        NetOffice.PowerPointApi.ChartGroup SurfaceGroup { get; }

        /// <summary>
        /// SupportByVersion PowerPoint 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff745066.aspx </remarks>
        [SupportByVersion("PowerPoint", 14, 15, 16)]
        NetOffice.PowerPointApi.Application Application { get; }

        /// <summary>
        /// SupportByVersion PowerPoint 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff744513.aspx </remarks>
        [SupportByVersion("PowerPoint", 14, 15, 16)]
        bool AutoScaling { get; set; }

        /// <summary>
        /// SupportByVersion PowerPoint 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff744327.aspx </remarks>
        [SupportByVersion("PowerPoint", 14, 15, 16)]
        NetOffice.PowerPointApi.ChartArea ChartArea { get; }

        /// <summary>
        /// SupportByVersion PowerPoint 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff743961.aspx </remarks>
        [SupportByVersion("PowerPoint", 14, 15, 16)]
        NetOffice.PowerPointApi.ChartTitle ChartTitle { get; }

        /// <summary>
        /// SupportByVersion PowerPoint 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("PowerPoint", 14, 15, 16)]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        NetOffice.PowerPointApi.Corners Corners { get; }

        /// <summary>
        /// SupportByVersion PowerPoint 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff746755.aspx </remarks>
        [SupportByVersion("PowerPoint", 14, 15, 16)]
        Int32 DepthPercent { get; set; }

        /// <summary>
        /// SupportByVersion PowerPoint 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff745600.aspx </remarks>
        [SupportByVersion("PowerPoint", 14, 15, 16)]
        NetOffice.PowerPointApi.Enums.XlDisplayBlanksAs DisplayBlanksAs { get; set; }

        /// <summary>
        /// SupportByVersion PowerPoint 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff745750.aspx </remarks>
        [SupportByVersion("PowerPoint", 14, 15, 16)]
        Int32 Elevation { get; set; }

        /// <summary>
        /// SupportByVersion PowerPoint 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff745846.aspx </remarks>
        [SupportByVersion("PowerPoint", 14, 15, 16)]
        NetOffice.PowerPointApi.Floor Floor { get; }

        /// <summary>
        /// SupportByVersion PowerPoint 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff746511.aspx </remarks>
        [SupportByVersion("PowerPoint", 14, 15, 16)]
        Int32 GapDepth { get; set; }

        /// <summary>
        /// SupportByVersion PowerPoint 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff746650.aspx </remarks>
        [SupportByVersion("PowerPoint", 14, 15, 16)]
        new object HasAxis { get; set; }

        /// <summary>
        /// SupportByVersion PowerPoint 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff743935.aspx </remarks>
        [SupportByVersion("PowerPoint", 14, 15, 16)]
        bool HasLegend { get; set; }

        /// <summary>
        /// SupportByVersion PowerPoint 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff746534.aspx </remarks>
        [SupportByVersion("PowerPoint", 14, 15, 16)]
        bool HasTitle { get; set; }

        /// <summary>
        /// SupportByVersion PowerPoint 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff745241.aspx </remarks>
        [SupportByVersion("PowerPoint", 14, 15, 16)]
        Int32 HeightPercent { get; set; }

        /// <summary>
        /// SupportByVersion PowerPoint 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff744151.aspx </remarks>
        [SupportByVersion("PowerPoint", 14, 15, 16)]
        NetOffice.PowerPointApi.Legend Legend { get; }

        /// <summary>
        /// SupportByVersion PowerPoint 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff744105.aspx </remarks>
        [SupportByVersion("PowerPoint", 14, 15, 16)]
        string Name { get; set; }

        /// <summary>
        /// SupportByVersion PowerPoint 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff743957.aspx </remarks>
        [SupportByVersion("PowerPoint", 14, 15, 16)]
        Int32 Perspective { get; set; }

        /// <summary>
        /// SupportByVersion PowerPoint 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff746093.aspx </remarks>
        [SupportByVersion("PowerPoint", 14, 15, 16)]
        NetOffice.PowerPointApi.PlotArea PlotArea { get; }

        /// <summary>
        /// SupportByVersion PowerPoint 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff745749.aspx </remarks>
        [SupportByVersion("PowerPoint", 14, 15, 16)]
        bool PlotVisibleOnly { get; set; }

        /// <summary>
        /// SupportByVersion PowerPoint 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff744814.aspx </remarks>
        [SupportByVersion("PowerPoint", 14, 15, 16)]
        object RightAngleAxes { get; set; }

        /// <summary>
        /// SupportByVersion PowerPoint 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff745024.aspx </remarks>
        [SupportByVersion("PowerPoint", 14, 15, 16)]
        object Rotation { get; set; }

        /// <summary>
        /// SupportByVersion PowerPoint 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("PowerPoint", 14, 15, 16)]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        Int32 Subtype { get; set; }

        /// <summary>
        /// SupportByVersion PowerPoint 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("PowerPoint", 14, 15, 16)]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        Int32 Type { get; set; }

        /// <summary>
        /// SupportByVersion PowerPoint 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff746542.aspx </remarks>
        [SupportByVersion("PowerPoint", 14, 15, 16)]
        NetOffice.PowerPointApi.Walls Walls { get; }

        /// <summary>
        /// SupportByVersion PowerPoint 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff745294.aspx </remarks>
        [SupportByVersion("PowerPoint", 14, 15, 16)]
        NetOffice.PowerPointApi.ChartFormat Format { get; }

        /// <summary>
        /// SupportByVersion PowerPoint 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff745821.aspx </remarks>
        [SupportByVersion("PowerPoint", 14, 15, 16)]
        bool ShowReportFilterFieldButtons { get; set; }

        /// <summary>
        /// SupportByVersion PowerPoint 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff743877.aspx </remarks>
        [SupportByVersion("PowerPoint", 14, 15, 16)]
        bool ShowLegendFieldButtons { get; set; }

        /// <summary>
        /// SupportByVersion PowerPoint 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff744539.aspx </remarks>
        [SupportByVersion("PowerPoint", 14, 15, 16)]
        bool ShowAxisFieldButtons { get; set; }

        /// <summary>
        /// SupportByVersion PowerPoint 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff746204.aspx </remarks>
        [SupportByVersion("PowerPoint", 14, 15, 16)]
        bool ShowValueFieldButtons { get; set; }

        /// <summary>
        /// SupportByVersion PowerPoint 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff744868.aspx </remarks>
        [SupportByVersion("PowerPoint", 14, 15, 16)]
        bool ShowAllFieldButtons { get; set; }

        /// <summary>
        /// SupportByVersion PowerPoint 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff746125.aspx </remarks>
        [SupportByVersion("PowerPoint", 14, 15, 16)]
        string AlternativeText { get; set; }

        /// <summary>
        /// SupportByVersion PowerPoint 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff745833.aspx </remarks>
        [SupportByVersion("PowerPoint", 14, 15, 16)]
        string Title { get; set; }

        /// <summary>
        /// SupportByVersion PowerPoint 15,16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/jj229264.aspx </remarks>
        [SupportByVersion("PowerPoint", 15, 16)]
        NetOffice.PowerPointApi.Enums.XlCategoryLabelLevel CategoryLabelLevel { get; set; }

        /// <summary>
        /// SupportByVersion PowerPoint 15,16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/jj228519.aspx </remarks>
        [SupportByVersion("PowerPoint", 15, 16)]
        NetOffice.PowerPointApi.Enums.XlSeriesNameLevel SeriesNameLevel { get; set; }

        /// <summary>
        /// SupportByVersion PowerPoint 15,16
        /// Get
        /// </summary>
        [SupportByVersion("PowerPoint", 15, 16)]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        bool HasHiddenContent { get; }

        /// <summary>
        /// SupportByVersion PowerPoint 15,16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/jj230443.aspx </remarks>
        [SupportByVersion("PowerPoint", 15, 16)]
        object ChartColor { get; set; }

        #endregion

        #region Methods

        /// <summary>
        /// SupportByVersion PowerPoint 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff745773.aspx </remarks>
        /// <param name="type">optional NetOffice.PowerPointApi.Enums.XlDataLabelsType Type = 2</param>
        /// <param name="legendKey">optional object legendKey</param>
        /// <param name="autoText">optional object autoText</param>
        /// <param name="hasLeaderLines">optional object hasLeaderLines</param>
        /// <param name="showSeriesName">optional object showSeriesName</param>
        /// <param name="showCategoryName">optional object showCategoryName</param>
        /// <param name="showValue">optional object showValue</param>
        /// <param name="showPercentage">optional object showPercentage</param>
        /// <param name="showBubbleSize">optional object showBubbleSize</param>
        /// <param name="separator">optional object separator</param>
        [SupportByVersion("PowerPoint", 14, 15, 16)]
        void ApplyDataLabels(object type, object legendKey, object autoText, object hasLeaderLines, object showSeriesName, object showCategoryName, object showValue, object showPercentage, object showBubbleSize, object separator);

        /// <summary>
        /// SupportByVersion PowerPoint 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff745773.aspx </remarks>
        [CustomMethod]
        [SupportByVersion("PowerPoint", 14, 15, 16)]
        void ApplyDataLabels();

        /// <summary>
        /// SupportByVersion PowerPoint 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff745773.aspx </remarks>
        /// <param name="type">optional NetOffice.PowerPointApi.Enums.XlDataLabelsType Type = 2</param>
        [CustomMethod]
        [SupportByVersion("PowerPoint", 14, 15, 16)]
        void ApplyDataLabels(object type);

        /// <summary>
        /// SupportByVersion PowerPoint 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff745773.aspx </remarks>
        /// <param name="type">optional NetOffice.PowerPointApi.Enums.XlDataLabelsType Type = 2</param>
        /// <param name="legendKey">optional object legendKey</param>
        [CustomMethod]
        [SupportByVersion("PowerPoint", 14, 15, 16)]
        void ApplyDataLabels(object type, object legendKey);

        /// <summary>
        /// SupportByVersion PowerPoint 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff745773.aspx </remarks>
        /// <param name="type">optional NetOffice.PowerPointApi.Enums.XlDataLabelsType Type = 2</param>
        /// <param name="legendKey">optional object legendKey</param>
        /// <param name="autoText">optional object autoText</param>
        [CustomMethod]
        [SupportByVersion("PowerPoint", 14, 15, 16)]
        void ApplyDataLabels(object type, object legendKey, object autoText);

        /// <summary>
        /// SupportByVersion PowerPoint 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff745773.aspx </remarks>
        /// <param name="type">optional NetOffice.PowerPointApi.Enums.XlDataLabelsType Type = 2</param>
        /// <param name="legendKey">optional object legendKey</param>
        /// <param name="autoText">optional object autoText</param>
        /// <param name="hasLeaderLines">optional object hasLeaderLines</param>
        [CustomMethod]
        [SupportByVersion("PowerPoint", 14, 15, 16)]
        void ApplyDataLabels(object type, object legendKey, object autoText, object hasLeaderLines);

        /// <summary>
        /// SupportByVersion PowerPoint 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff745773.aspx </remarks>
        /// <param name="type">optional NetOffice.PowerPointApi.Enums.XlDataLabelsType Type = 2</param>
        /// <param name="legendKey">optional object legendKey</param>
        /// <param name="autoText">optional object autoText</param>
        /// <param name="hasLeaderLines">optional object hasLeaderLines</param>
        /// <param name="showSeriesName">optional object showSeriesName</param>
        [CustomMethod]
        [SupportByVersion("PowerPoint", 14, 15, 16)]
        void ApplyDataLabels(object type, object legendKey, object autoText, object hasLeaderLines, object showSeriesName);

        /// <summary>
        /// SupportByVersion PowerPoint 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff745773.aspx </remarks>
        /// <param name="type">optional NetOffice.PowerPointApi.Enums.XlDataLabelsType Type = 2</param>
        /// <param name="legendKey">optional object legendKey</param>
        /// <param name="autoText">optional object autoText</param>
        /// <param name="hasLeaderLines">optional object hasLeaderLines</param>
        /// <param name="showSeriesName">optional object showSeriesName</param>
        /// <param name="showCategoryName">optional object showCategoryName</param>
        [CustomMethod]
        [SupportByVersion("PowerPoint", 14, 15, 16)]
        void ApplyDataLabels(object type, object legendKey, object autoText, object hasLeaderLines, object showSeriesName, object showCategoryName);

        /// <summary>
        /// SupportByVersion PowerPoint 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff745773.aspx </remarks>
        /// <param name="type">optional NetOffice.PowerPointApi.Enums.XlDataLabelsType Type = 2</param>
        /// <param name="legendKey">optional object legendKey</param>
        /// <param name="autoText">optional object autoText</param>
        /// <param name="hasLeaderLines">optional object hasLeaderLines</param>
        /// <param name="showSeriesName">optional object showSeriesName</param>
        /// <param name="showCategoryName">optional object showCategoryName</param>
        /// <param name="showValue">optional object showValue</param>
        [CustomMethod]
        [SupportByVersion("PowerPoint", 14, 15, 16)]
        void ApplyDataLabels(object type, object legendKey, object autoText, object hasLeaderLines, object showSeriesName, object showCategoryName, object showValue);

        /// <summary>
        /// SupportByVersion PowerPoint 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff745773.aspx </remarks>
        /// <param name="type">optional NetOffice.PowerPointApi.Enums.XlDataLabelsType Type = 2</param>
        /// <param name="legendKey">optional object legendKey</param>
        /// <param name="autoText">optional object autoText</param>
        /// <param name="hasLeaderLines">optional object hasLeaderLines</param>
        /// <param name="showSeriesName">optional object showSeriesName</param>
        /// <param name="showCategoryName">optional object showCategoryName</param>
        /// <param name="showValue">optional object showValue</param>
        /// <param name="showPercentage">optional object showPercentage</param>
        [CustomMethod]
        [SupportByVersion("PowerPoint", 14, 15, 16)]
        void ApplyDataLabels(object type, object legendKey, object autoText, object hasLeaderLines, object showSeriesName, object showCategoryName, object showValue, object showPercentage);

        /// <summary>
        /// SupportByVersion PowerPoint 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff745773.aspx </remarks>
        /// <param name="type">optional NetOffice.PowerPointApi.Enums.XlDataLabelsType Type = 2</param>
        /// <param name="legendKey">optional object legendKey</param>
        /// <param name="autoText">optional object autoText</param>
        /// <param name="hasLeaderLines">optional object hasLeaderLines</param>
        /// <param name="showSeriesName">optional object showSeriesName</param>
        /// <param name="showCategoryName">optional object showCategoryName</param>
        /// <param name="showValue">optional object showValue</param>
        /// <param name="showPercentage">optional object showPercentage</param>
        /// <param name="showBubbleSize">optional object showBubbleSize</param>
        [CustomMethod]
        [SupportByVersion("PowerPoint", 14, 15, 16)]
        void ApplyDataLabels(object type, object legendKey, object autoText, object hasLeaderLines, object showSeriesName, object showCategoryName, object showValue, object showPercentage, object showBubbleSize);

        /// <summary>
        /// SupportByVersion PowerPoint 14, 15, 16
        /// </summary>
        /// <param name="chartType">NetOffice.OfficeApi.Enums.XlChartType chartType</param>
        /// <param name="typeName">optional object typeName</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [SupportByVersion("PowerPoint", 14, 15, 16)]
        void ApplyCustomType(NetOffice.OfficeApi.Enums.XlChartType chartType, object typeName);

        /// <summary>
        /// SupportByVersion PowerPoint 14, 15, 16
        /// </summary>
        /// <param name="chartType">NetOffice.OfficeApi.Enums.XlChartType chartType</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [CustomMethod]
        [SupportByVersion("PowerPoint", 14, 15, 16)]
        void ApplyCustomType(NetOffice.OfficeApi.Enums.XlChartType chartType);

        /// <summary>
        /// SupportByVersion PowerPoint 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff746151.aspx </remarks>
        /// <param name="x">Int32 x</param>
        /// <param name="y">Int32 y</param>
        /// <param name="elementID">Int32 elementID</param>
        /// <param name="arg1">Int32 arg1</param>
        /// <param name="arg2">Int32 arg2</param>
        [SupportByVersion("PowerPoint", 14, 15, 16)]
        void GetChartElement(Int32 x, Int32 y, Int32 elementID, Int32 arg1, Int32 arg2);

        /// <summary>
        /// SupportByVersion PowerPoint 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff746759.aspx </remarks>
        /// <param name="source">string source</param>
        /// <param name="plotBy">optional object plotBy</param>
        [SupportByVersion("PowerPoint", 14, 15, 16)]
        void SetSourceData(string source, object plotBy);

        /// <summary>
        /// SupportByVersion PowerPoint 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff746759.aspx </remarks>
        /// <param name="source">string source</param>
        [CustomMethod]
        [SupportByVersion("PowerPoint", 14, 15, 16)]
        void SetSourceData(string source);

        /// <summary>
        /// SupportByVersion PowerPoint 14, 15, 16
        /// </summary>
        /// <param name="gallery">Int32 gallery</param>
        /// <param name="format">optional object format</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [SupportByVersion("PowerPoint", 14, 15, 16)]
        void AutoFormat(Int32 gallery, object format);

        /// <summary>
        /// SupportByVersion PowerPoint 14, 15, 16
        /// </summary>
        /// <param name="gallery">Int32 gallery</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [CustomMethod]
        [SupportByVersion("PowerPoint", 14, 15, 16)]
        void AutoFormat(Int32 gallery);

        /// <summary>
        /// SupportByVersion PowerPoint 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff745424.aspx </remarks>
        /// <param name="fileName">string fileName</param>
        [SupportByVersion("PowerPoint", 14, 15, 16)]
        void SetBackgroundPicture(string fileName);

        /// <summary>
        /// SupportByVersion PowerPoint 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff746056.aspx </remarks>
        /// <param name="type">optional object type</param>
        [SupportByVersion("PowerPoint", 14, 15, 16)]
        void Paste(object type);

        /// <summary>
        /// SupportByVersion PowerPoint 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff746056.aspx </remarks>
        [CustomMethod]
        [SupportByVersion("PowerPoint", 14, 15, 16)]
        void Paste();

        /// <summary>
        /// SupportByVersion PowerPoint 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff745864.aspx </remarks>
        /// <param name="name">object name</param>
        [SupportByVersion("PowerPoint", 14, 15, 16)]
        void SetDefaultChart(object name);

        /// <summary>
        /// SupportByVersion PowerPoint 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff744899.aspx </remarks>
        /// <param name="fileName">string fileName</param>
        [SupportByVersion("PowerPoint", 14, 15, 16)]
        void ApplyChartTemplate(string fileName);

        /// <summary>
        /// SupportByVersion PowerPoint 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff744919.aspx </remarks>
        /// <param name="fileName">string fileName</param>
        [SupportByVersion("PowerPoint", 14, 15, 16)]
        void SaveChartTemplate(string fileName);

        /// <summary>
        /// SupportByVersion PowerPoint 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff746785.aspx </remarks>
        [SupportByVersion("PowerPoint", 14, 15, 16)]
        void ClearToMatchStyle();

        /// <summary>
        /// SupportByVersion PowerPoint 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff745663.aspx </remarks>
        /// <param name="layout">Int32 layout</param>
        /// <param name="chartType">optional object chartType</param>
        [SupportByVersion("PowerPoint", 14, 15, 16)]
        void ApplyLayout(Int32 layout, object chartType);

        /// <summary>
        /// SupportByVersion PowerPoint 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff745663.aspx </remarks>
        /// <param name="layout">Int32 layout</param>
        [CustomMethod]
        [SupportByVersion("PowerPoint", 14, 15, 16)]
        void ApplyLayout(Int32 layout);

        /// <summary>
        /// SupportByVersion PowerPoint 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff745006.aspx </remarks>
        [SupportByVersion("PowerPoint", 14, 15, 16)]
        void Refresh();

        /// <summary>
        /// SupportByVersion PowerPoint 14, 15, 16
        /// </summary>
        /// <param name="index">optional object index</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [SupportByVersion("PowerPoint", 14, 15, 16)]
        object AreaGroups(object index);

        /// <summary>
        /// SupportByVersion PowerPoint 14, 15, 16
        /// </summary>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [CustomMethod]
        [SupportByVersion("PowerPoint", 14, 15, 16)]
        object AreaGroups();

        /// <summary>
        /// SupportByVersion PowerPoint 14, 15, 16
        /// </summary>
        /// <param name="index">optional object index</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [SupportByVersion("PowerPoint", 14, 15, 16)]
        object BarGroups(object index);

        /// <summary>
        /// SupportByVersion PowerPoint 14, 15, 16
        /// </summary>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [CustomMethod]
        [SupportByVersion("PowerPoint", 14, 15, 16)]
        object BarGroups();

        /// <summary>
        /// SupportByVersion PowerPoint 14, 15, 16
        /// </summary>
        /// <param name="index">optional object index</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [SupportByVersion("PowerPoint", 14, 15, 16)]
        object ColumnGroups(object index);

        /// <summary>
        /// SupportByVersion PowerPoint 14, 15, 16
        /// </summary>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [CustomMethod]
        [SupportByVersion("PowerPoint", 14, 15, 16)]
        object ColumnGroups();

        /// <summary>
        /// SupportByVersion PowerPoint 14, 15, 16
        /// </summary>
        /// <param name="index">optional object index</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [SupportByVersion("PowerPoint", 14, 15, 16)]
        object LineGroups(object index);

        /// <summary>
        /// SupportByVersion PowerPoint 14, 15, 16
        /// </summary>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [CustomMethod]
        [SupportByVersion("PowerPoint", 14, 15, 16)]
        object LineGroups();

        /// <summary>
        /// SupportByVersion PowerPoint 14, 15, 16
        /// </summary>
        /// <param name="index">optional object index</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [SupportByVersion("PowerPoint", 14, 15, 16)]
        object PieGroups(object index);

        /// <summary>
        /// SupportByVersion PowerPoint 14, 15, 16
        /// </summary>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [CustomMethod]
        [SupportByVersion("PowerPoint", 14, 15, 16)]
        object PieGroups();

        /// <summary>
        /// SupportByVersion PowerPoint 14, 15, 16
        /// </summary>
        /// <param name="index">optional object index</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [SupportByVersion("PowerPoint", 14, 15, 16)]
        object DoughnutGroups(object index);

        /// <summary>
        /// SupportByVersion PowerPoint 14, 15, 16
        /// </summary>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [CustomMethod]
        [SupportByVersion("PowerPoint", 14, 15, 16)]
        object DoughnutGroups();

        /// <summary>
        /// SupportByVersion PowerPoint 14, 15, 16
        /// </summary>
        /// <param name="index">optional object index</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [SupportByVersion("PowerPoint", 14, 15, 16)]
        object RadarGroups(object index);

        /// <summary>
        /// SupportByVersion PowerPoint 14, 15, 16
        /// </summary>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [CustomMethod]
        [SupportByVersion("PowerPoint", 14, 15, 16)]
        object RadarGroups();

        /// <summary>
        /// SupportByVersion PowerPoint 14, 15, 16
        /// </summary>
        /// <param name="index">optional object index</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [SupportByVersion("PowerPoint", 14, 15, 16)]
        object XYGroups(object index);

        /// <summary>
        /// SupportByVersion PowerPoint 14, 15, 16
        /// </summary>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [CustomMethod]
        [SupportByVersion("PowerPoint", 14, 15, 16)]
        object XYGroups();

        /// <summary>
        /// SupportByVersion PowerPoint 14, 15, 16
        /// </summary>
        /// <param name="type">optional NetOffice.PowerPointApi.Enums.XlDataLabelsType Type = 2</param>
        /// <param name="legendKey">optional object legendKey</param>
        /// <param name="autoText">optional object autoText</param>
        /// <param name="hasLeaderLines">optional object hasLeaderLines</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [SupportByVersion("PowerPoint", 14, 15, 16)]
        void _ApplyDataLabels(object type, object legendKey, object autoText, object hasLeaderLines);

        /// <summary>
        /// SupportByVersion PowerPoint 14, 15, 16
        /// </summary>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [CustomMethod]
        [SupportByVersion("PowerPoint", 14, 15, 16)]
        void _ApplyDataLabels();

        /// <summary>
        /// SupportByVersion PowerPoint 14, 15, 16
        /// </summary>
        /// <param name="type">optional NetOffice.PowerPointApi.Enums.XlDataLabelsType Type = 2</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [CustomMethod]
        [SupportByVersion("PowerPoint", 14, 15, 16)]
        void _ApplyDataLabels(object type);

        /// <summary>
        /// SupportByVersion PowerPoint 14, 15, 16
        /// </summary>
        /// <param name="type">optional NetOffice.PowerPointApi.Enums.XlDataLabelsType Type = 2</param>
        /// <param name="legendKey">optional object legendKey</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [CustomMethod]
        [SupportByVersion("PowerPoint", 14, 15, 16)]
        void _ApplyDataLabels(object type, object legendKey);

        /// <summary>
        /// SupportByVersion PowerPoint 14, 15, 16
        /// </summary>
        /// <param name="type">optional NetOffice.PowerPointApi.Enums.XlDataLabelsType Type = 2</param>
        /// <param name="legendKey">optional object legendKey</param>
        /// <param name="autoText">optional object autoText</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [CustomMethod]
        [SupportByVersion("PowerPoint", 14, 15, 16)]
        void _ApplyDataLabels(object type, object legendKey, object autoText);

        /// <summary>
        /// SupportByVersion PowerPoint 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff745216.aspx </remarks>
        /// <param name="type">optional object type</param>
        /// <param name="axisGroup">optional NetOffice.PowerPointApi.Enums.XlAxisGroup AxisGroup = 1</param>
        [SupportByVersion("PowerPoint", 14, 15, 16)]
        object Axes(object type, object axisGroup);

        /// <summary>
        /// SupportByVersion PowerPoint 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff745216.aspx </remarks>
        [CustomMethod]
        [SupportByVersion("PowerPoint", 14, 15, 16)]
        object Axes();

        /// <summary>
        /// SupportByVersion PowerPoint 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff745216.aspx </remarks>
        /// <param name="type">optional object type</param>
        [CustomMethod]
        [SupportByVersion("PowerPoint", 14, 15, 16)]
        object Axes(object type);

        /// <summary>
        /// SupportByVersion PowerPoint 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff744238.aspx </remarks>
        /// <param name="index">optional object index</param>
        [SupportByVersion("PowerPoint", 14, 15, 16)]
        object ChartGroups(object index);

        /// <summary>
        /// SupportByVersion PowerPoint 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff744238.aspx </remarks>
        [CustomMethod]
        [SupportByVersion("PowerPoint", 14, 15, 16)]
        object ChartGroups();

        /// <summary>
        /// SupportByVersion PowerPoint 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff745899.aspx </remarks>
        /// <param name="source">optional object source</param>
        /// <param name="gallery">optional object gallery</param>
        /// <param name="format">optional object format</param>
        /// <param name="plotBy">optional object plotBy</param>
        /// <param name="categoryLabels">optional object categoryLabels</param>
        /// <param name="seriesLabels">optional object seriesLabels</param>
        /// <param name="hasLegend">optional object hasLegend</param>
        /// <param name="title">optional object title</param>
        /// <param name="categoryTitle">optional object categoryTitle</param>
        /// <param name="valueTitle">optional object valueTitle</param>
        /// <param name="extraTitle">optional object extraTitle</param>
        [SupportByVersion("PowerPoint", 14, 15, 16)]
        void ChartWizard(object source, object gallery, object format, object plotBy, object categoryLabels, object seriesLabels, object hasLegend, object title, object categoryTitle, object valueTitle, object extraTitle);

        /// <summary>
        /// SupportByVersion PowerPoint 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff745899.aspx </remarks>
        [CustomMethod]
        [SupportByVersion("PowerPoint", 14, 15, 16)]
        void ChartWizard();

        /// <summary>
        /// SupportByVersion PowerPoint 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff745899.aspx </remarks>
        /// <param name="source">optional object source</param>
        [CustomMethod]
        [SupportByVersion("PowerPoint", 14, 15, 16)]
        void ChartWizard(object source);

        /// <summary>
        /// SupportByVersion PowerPoint 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff745899.aspx </remarks>
        /// <param name="source">optional object source</param>
        /// <param name="gallery">optional object gallery</param>
        [CustomMethod]
        [SupportByVersion("PowerPoint", 14, 15, 16)]
        void ChartWizard(object source, object gallery);

        /// <summary>
        /// SupportByVersion PowerPoint 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff745899.aspx </remarks>
        /// <param name="source">optional object source</param>
        /// <param name="gallery">optional object gallery</param>
        /// <param name="format">optional object format</param>
        [CustomMethod]
        [SupportByVersion("PowerPoint", 14, 15, 16)]
        void ChartWizard(object source, object gallery, object format);

        /// <summary>
        /// SupportByVersion PowerPoint 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff745899.aspx </remarks>
        /// <param name="source">optional object source</param>
        /// <param name="gallery">optional object gallery</param>
        /// <param name="format">optional object format</param>
        /// <param name="plotBy">optional object plotBy</param>
        [CustomMethod]
        [SupportByVersion("PowerPoint", 14, 15, 16)]
        void ChartWizard(object source, object gallery, object format, object plotBy);

        /// <summary>
        /// SupportByVersion PowerPoint 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff745899.aspx </remarks>
        /// <param name="source">optional object source</param>
        /// <param name="gallery">optional object gallery</param>
        /// <param name="format">optional object format</param>
        /// <param name="plotBy">optional object plotBy</param>
        /// <param name="categoryLabels">optional object categoryLabels</param>
        [CustomMethod]
        [SupportByVersion("PowerPoint", 14, 15, 16)]
        void ChartWizard(object source, object gallery, object format, object plotBy, object categoryLabels);

        /// <summary>
        /// SupportByVersion PowerPoint 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff745899.aspx </remarks>
        /// <param name="source">optional object source</param>
        /// <param name="gallery">optional object gallery</param>
        /// <param name="format">optional object format</param>
        /// <param name="plotBy">optional object plotBy</param>
        /// <param name="categoryLabels">optional object categoryLabels</param>
        /// <param name="seriesLabels">optional object seriesLabels</param>
        [CustomMethod]
        [SupportByVersion("PowerPoint", 14, 15, 16)]
        void ChartWizard(object source, object gallery, object format, object plotBy, object categoryLabels, object seriesLabels);

        /// <summary>
        /// SupportByVersion PowerPoint 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff745899.aspx </remarks>
        /// <param name="source">optional object source</param>
        /// <param name="gallery">optional object gallery</param>
        /// <param name="format">optional object format</param>
        /// <param name="plotBy">optional object plotBy</param>
        /// <param name="categoryLabels">optional object categoryLabels</param>
        /// <param name="seriesLabels">optional object seriesLabels</param>
        /// <param name="hasLegend">optional object hasLegend</param>
        [CustomMethod]
        [SupportByVersion("PowerPoint", 14, 15, 16)]
        void ChartWizard(object source, object gallery, object format, object plotBy, object categoryLabels, object seriesLabels, object hasLegend);

        /// <summary>
        /// SupportByVersion PowerPoint 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff745899.aspx </remarks>
        /// <param name="source">optional object source</param>
        /// <param name="gallery">optional object gallery</param>
        /// <param name="format">optional object format</param>
        /// <param name="plotBy">optional object plotBy</param>
        /// <param name="categoryLabels">optional object categoryLabels</param>
        /// <param name="seriesLabels">optional object seriesLabels</param>
        /// <param name="hasLegend">optional object hasLegend</param>
        /// <param name="title">optional object title</param>
        [CustomMethod]
        [SupportByVersion("PowerPoint", 14, 15, 16)]
        void ChartWizard(object source, object gallery, object format, object plotBy, object categoryLabels, object seriesLabels, object hasLegend, object title);

        /// <summary>
        /// SupportByVersion PowerPoint 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff745899.aspx </remarks>
        /// <param name="source">optional object source</param>
        /// <param name="gallery">optional object gallery</param>
        /// <param name="format">optional object format</param>
        /// <param name="plotBy">optional object plotBy</param>
        /// <param name="categoryLabels">optional object categoryLabels</param>
        /// <param name="seriesLabels">optional object seriesLabels</param>
        /// <param name="hasLegend">optional object hasLegend</param>
        /// <param name="title">optional object title</param>
        /// <param name="categoryTitle">optional object categoryTitle</param>
        [CustomMethod]
        [SupportByVersion("PowerPoint", 14, 15, 16)]
        void ChartWizard(object source, object gallery, object format, object plotBy, object categoryLabels, object seriesLabels, object hasLegend, object title, object categoryTitle);

        /// <summary>
        /// SupportByVersion PowerPoint 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff745899.aspx </remarks>
        /// <param name="source">optional object source</param>
        /// <param name="gallery">optional object gallery</param>
        /// <param name="format">optional object format</param>
        /// <param name="plotBy">optional object plotBy</param>
        /// <param name="categoryLabels">optional object categoryLabels</param>
        /// <param name="seriesLabels">optional object seriesLabels</param>
        /// <param name="hasLegend">optional object hasLegend</param>
        /// <param name="title">optional object title</param>
        /// <param name="categoryTitle">optional object categoryTitle</param>
        /// <param name="valueTitle">optional object valueTitle</param>
        [CustomMethod]
        [SupportByVersion("PowerPoint", 14, 15, 16)]
        void ChartWizard(object source, object gallery, object format, object plotBy, object categoryLabels, object seriesLabels, object hasLegend, object title, object categoryTitle, object valueTitle);

        /// <summary>
        /// SupportByVersion PowerPoint 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff745934.aspx </remarks>
        /// <param name="before">optional object before</param>
        /// <param name="after">optional object after</param>
        [SupportByVersion("PowerPoint", 14, 15, 16)]
        void Copy(object before, object after);

        /// <summary>
        /// SupportByVersion PowerPoint 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff745934.aspx </remarks>
        [CustomMethod]
        [SupportByVersion("PowerPoint", 14, 15, 16)]
        void Copy();

        /// <summary>
        /// SupportByVersion PowerPoint 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff745934.aspx </remarks>
        /// <param name="before">optional object before</param>
        [CustomMethod]
        [SupportByVersion("PowerPoint", 14, 15, 16)]
        void Copy(object before);

        /// <summary>
        /// SupportByVersion PowerPoint 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff745919.aspx </remarks>
        /// <param name="appearance">optional NetOffice.PowerPointApi.Enums.XlPictureAppearance Appearance = 1</param>
        /// <param name="format">optional NetOffice.PowerPointApi.Enums.XlCopyPictureFormat Format = -4147</param>
        /// <param name="size">optional NetOffice.PowerPointApi.Enums.XlPictureAppearance Size = 2</param>
        [SupportByVersion("PowerPoint", 14, 15, 16)]
        void CopyPicture(object appearance, object format, object size);

        /// <summary>
        /// SupportByVersion PowerPoint 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff745919.aspx </remarks>
        [CustomMethod]
        [SupportByVersion("PowerPoint", 14, 15, 16)]
        void CopyPicture();

        /// <summary>
        /// SupportByVersion PowerPoint 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff745919.aspx </remarks>
        /// <param name="appearance">optional NetOffice.PowerPointApi.Enums.XlPictureAppearance Appearance = 1</param>
        [CustomMethod]
        [SupportByVersion("PowerPoint", 14, 15, 16)]
        void CopyPicture(object appearance);

        /// <summary>
        /// SupportByVersion PowerPoint 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff745919.aspx </remarks>
        /// <param name="appearance">optional NetOffice.PowerPointApi.Enums.XlPictureAppearance Appearance = 1</param>
        /// <param name="format">optional NetOffice.PowerPointApi.Enums.XlCopyPictureFormat Format = -4147</param>
        [CustomMethod]
        [SupportByVersion("PowerPoint", 14, 15, 16)]
        void CopyPicture(object appearance, object format);

        /// <summary>
        /// SupportByVersion PowerPoint 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff745109.aspx </remarks>
        [SupportByVersion("PowerPoint", 14, 15, 16)]
        void Delete();

        /// <summary>
        /// SupportByVersion PowerPoint 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff744128.aspx </remarks>
        /// <param name="fileName">string fileName</param>
        /// <param name="filterName">optional object filterName</param>
        /// <param name="interactive">optional object interactive</param>
        [SupportByVersion("PowerPoint", 14, 15, 16)]
        bool Export(string fileName, object filterName, object interactive);

        /// <summary>
        /// SupportByVersion PowerPoint 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff744128.aspx </remarks>
        /// <param name="fileName">string fileName</param>
        [CustomMethod]
        [SupportByVersion("PowerPoint", 14, 15, 16)]
        bool Export(string fileName);

        /// <summary>
        /// SupportByVersion PowerPoint 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff744128.aspx </remarks>
        /// <param name="fileName">string fileName</param>
        /// <param name="filterName">optional object filterName</param>
        [CustomMethod]
        [SupportByVersion("PowerPoint", 14, 15, 16)]
        bool Export(string fileName, object filterName);

        /// <summary>
        /// SupportByVersion PowerPoint 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff745013.aspx </remarks>
        /// <param name="replace">optional object replace</param>
        [SupportByVersion("PowerPoint", 14, 15, 16)]
        void Select(object replace);

        /// <summary>
        /// SupportByVersion PowerPoint 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff745013.aspx </remarks>
        [CustomMethod]
        [SupportByVersion("PowerPoint", 14, 15, 16)]
        void Select();

        /// <summary>
        /// SupportByVersion PowerPoint 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff745538.aspx </remarks>
        /// <param name="index">optional object index</param>
        [SupportByVersion("PowerPoint", 14, 15, 16)]
        object SeriesCollection(object index);

        /// <summary>
        /// SupportByVersion PowerPoint 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff745538.aspx </remarks>
        [CustomMethod]
        [SupportByVersion("PowerPoint", 14, 15, 16)]
        object SeriesCollection();

        /// <summary>
        /// SupportByVersion PowerPoint 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff746262.aspx </remarks>
        /// <param name="element">NetOffice.OfficeApi.Enums.MsoChartElementType element</param>
        [SupportByVersion("PowerPoint", 14, 15, 16)]
        void SetElement(NetOffice.OfficeApi.Enums.MsoChartElementType element);

        /// <summary>
        /// SupportByVersion PowerPoint 15,16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/jj228028.aspx </remarks>
        /// <param name="index">optional object index</param>
        [SupportByVersion("PowerPoint", 15, 16)]
        object FullSeriesCollection(object index);

        /// <summary>
        /// SupportByVersion PowerPoint 15,16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/jj228028.aspx </remarks>
        [CustomMethod]
        [SupportByVersion("PowerPoint", 15, 16)]
        object FullSeriesCollection();

        /// <summary>
        /// SupportByVersion PowerPoint 15,16
        /// </summary>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [SupportByVersion("PowerPoint", 15, 16)]
        void DeleteHiddenContent();

        /// <summary>
        /// SupportByVersion PowerPoint 15,16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/jj227229.aspx </remarks>
        [SupportByVersion("PowerPoint", 15, 16)]
        void ClearToMatchColorStyle();

        #endregion
    }  
}
