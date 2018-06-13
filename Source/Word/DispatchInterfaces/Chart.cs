using System.Reflection;
using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.WordApi
{
    /// <summary>
    /// Chart
    /// </summary>
    [SyntaxBypass]
    public interface Chart_ : ICOMObject
    {
        #region Properties

        /// <summary>
        /// SupportByVersion Word 14, 15, 16
        /// Get
        /// Unknown COM Proxy
        /// </summary>
        /// <param name="index">optional object index</param>
        /// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff836391.aspx
        [SupportByVersion("Word", 14, 15, 16), ProxyResult]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        object get_ChartGroups(object index);

        /// <summary>
        /// SupportByVersion Word 14, 15, 16
        /// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff836391.aspx
        /// Alias for get_ChartGroups
        /// Unknown COM Proxy
        /// </summary>
        /// <param name="index">optional object index</param>
        [SupportByVersion("Word", 14, 15, 16), ProxyResult, Redirect("get_ChartGroups")]
        object ChartGroups(object index);

        /// <summary>
        /// SupportByVersion Word 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <param name="index1">optional object index1</param>
        /// <param name="index2">optional object index2</param>
        /// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff836998.aspx
        [SupportByVersion("Word", 14, 15, 16)]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        object get_HasAxis(object index1, object index2);

        /// <summary>
        /// SupportByVersion Word 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <param name="index1">optional object index1</param>
        /// <param name="index2">optional object index2</param>
        /// <param name="value">optional object value</param>
        [SupportByVersion("Word", 14, 15, 16)]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        void set_HasAxis(object index1, object index2, object value);

        /// <summary>
        /// SupportByVersion Word 14, 15, 16
        /// Alias for get_HasAxis
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff836998.aspx </remarks>
        /// <param name="index1">optional object index1</param>
        /// <param name="index2">optional object index2</param>
        [SupportByVersion("Word", 14, 15, 16), Redirect("get_HasAxis")]
        object HasAxis(object index1, object index2);

        /// <summary>
        /// SupportByVersion Word 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <param name="index1">optional object index1</param>
        /// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff836998.aspx
        [SupportByVersion("Word", 14, 15, 16)]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        object get_HasAxis(object index1);

        /// <summary>
        /// SupportByVersion Word 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <param name="index1">optional object index1</param>
        /// <param name="value">optional object value</param>
        [SupportByVersion("Word", 14, 15, 16)]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        void set_HasAxis(object index1, object value);

        /// <summary>
        /// SupportByVersion Word 14, 15, 16
        /// Alias for get_HasAxis
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff836998.aspx </remarks>
        /// <param name="index1">optional object index1</param>
        [SupportByVersion("Word", 14, 15, 16), Redirect("get_HasAxis")]
        object HasAxis(object index1);

        #endregion
    }

    /// <summary>
	/// DispatchInterface Chart 
	/// SupportByVersion Word, 14,15,16
	/// </summary>
	/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff193446.aspx </remarks>
	[SupportByVersion("Word", 14, 15, 16)]
    [EntityType(EntityType.IsDispatchInterface)]
	[TypeId("6FFA84BB-A350-4442-BB53-A43653459A84")]
    public interface Chart : Chart_
    {
        #region Properties

        /// <summary>
        /// SupportByVersion Word 14, 15, 16
        /// Get
        /// Unknown COM Proxy
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff191738.aspx </remarks>
        [SupportByVersion("Word", 14, 15, 16), ProxyResult]
        object Parent { get; }

        /// <summary>
        /// SupportByVersion Word 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff196350.aspx </remarks>
        [SupportByVersion("Word", 14, 15, 16)]
        bool HasTitle { get; set; }

        /// <summary>
        /// SupportByVersion Word 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff191751.aspx </remarks>
        [SupportByVersion("Word", 14, 15, 16)]
        NetOffice.WordApi.ChartTitle ChartTitle { get; }

        /// <summary>
        /// SupportByVersion Word 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff840907.aspx </remarks>
        [SupportByVersion("Word", 14, 15, 16)]
        Int32 DepthPercent { get; set; }

        /// <summary>
        /// SupportByVersion Word 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff192611.aspx </remarks>
        [SupportByVersion("Word", 14, 15, 16)]
        Int32 Elevation { get; set; }

        /// <summary>
        /// SupportByVersion Word 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff845244.aspx </remarks>
        [SupportByVersion("Word", 14, 15, 16)]
        Int32 GapDepth { get; set; }

        /// <summary>
        /// SupportByVersion Word 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff836594.aspx </remarks>
        [SupportByVersion("Word", 14, 15, 16)]
        Int32 HeightPercent { get; set; }

        /// <summary>
        /// SupportByVersion Word 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838954.aspx </remarks>
        [SupportByVersion("Word", 14, 15, 16)]
        Int32 Perspective { get; set; }

        /// <summary>
        /// SupportByVersion Word 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838938.aspx </remarks>
        [SupportByVersion("Word", 14, 15, 16)]
        object RightAngleAxes { get; set; }

        /// <summary>
        /// SupportByVersion Word 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff835465.aspx </remarks>
        [SupportByVersion("Word", 14, 15, 16)]
        object Rotation { get; set; }

        /// <summary>
        /// SupportByVersion Word 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff196216.aspx </remarks>
        [SupportByVersion("Word", 14, 15, 16)]
        NetOffice.WordApi.Enums.XlDisplayBlanksAs DisplayBlanksAs { get; set; }

        /// <summary>
        /// SupportByVersion Word 14, 15, 16
        /// Get
        /// Unknown COM Proxy
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff836391.aspx </remarks>
        [SupportByVersion("Word", 14, 15, 16), ProxyResult]
        new object ChartGroups { get; }

        /// <summary>
        /// SupportByVersion Word 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Word", 14, 15, 16)]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        Int32 SubType { get; set; }

        /// <summary>
        /// SupportByVersion Word 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Word", 14, 15, 16)]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        Int32 Type { get; set; }

        /// <summary>
        /// SupportByVersion Word 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Word", 14, 15, 16)]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        NetOffice.WordApi.Corners Corners { get; }

        /// <summary>
        /// SupportByVersion Word 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff836334.aspx </remarks>
        [SupportByVersion("Word", 14, 15, 16)]
        NetOffice.OfficeApi.Enums.XlChartType ChartType { get; set; }

        /// <summary>
        /// SupportByVersion Word 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff197158.aspx </remarks>
        [SupportByVersion("Word", 14, 15, 16)]
        bool HasDataTable { get; set; }

        /// <summary>
        /// SupportByVersion Word 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff836380.aspx </remarks>
        [SupportByVersion("Word", 14, 15, 16)]
        NetOffice.WordApi.Enums.XlRowCol PlotBy { get; set; }

        /// <summary>
        /// SupportByVersion Word 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff845054.aspx </remarks>
        [SupportByVersion("Word", 14, 15, 16)]
        bool HasLegend { get; set; }

        /// <summary>
        /// SupportByVersion Word 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff836685.aspx </remarks>
        [SupportByVersion("Word", 14, 15, 16)]
        NetOffice.WordApi.Legend Legend { get; }

        /// <summary>
        /// SupportByVersion Word 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff836998.aspx </remarks>
        [SupportByVersion("Word", 14, 15, 16)]
        new object HasAxis { get; set; }

        /// <summary>
        /// SupportByVersion Word 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff840511.aspx </remarks>
        [SupportByVersion("Word", 14, 15, 16)]
        NetOffice.WordApi.Walls Walls { get; }

        /// <summary>
        /// SupportByVersion Word 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff845855.aspx </remarks>
        [SupportByVersion("Word", 14, 15, 16)]
        NetOffice.WordApi.Floor Floor { get; }

        /// <summary>
        /// SupportByVersion Word 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194655.aspx </remarks>
        [SupportByVersion("Word", 14, 15, 16)]
        NetOffice.WordApi.PlotArea PlotArea { get; }

        /// <summary>
        /// SupportByVersion Word 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff196388.aspx </remarks>
        [SupportByVersion("Word", 14, 15, 16)]
        bool PlotVisibleOnly { get; set; }

        /// <summary>
        /// SupportByVersion Word 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff836658.aspx </remarks>
        [SupportByVersion("Word", 14, 15, 16)]
        NetOffice.WordApi.ChartArea ChartArea { get; }

        /// <summary>
        /// SupportByVersion Word 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff823268.aspx </remarks>
        [SupportByVersion("Word", 14, 15, 16)]
        bool AutoScaling { get; set; }

        /// <summary>
        /// SupportByVersion Word 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff191967.aspx </remarks>
        [SupportByVersion("Word", 14, 15, 16)]
        NetOffice.WordApi.DataTable DataTable { get; }

        /// <summary>
        /// SupportByVersion Word 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff839500.aspx </remarks>
        [SupportByVersion("Word", 14, 15, 16)]
        NetOffice.WordApi.Enums.XlBarShape BarShape { get; set; }

        /// <summary>
        /// SupportByVersion Word 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff839285.aspx </remarks>
        [SupportByVersion("Word", 14, 15, 16)]
        NetOffice.WordApi.Walls SideWall { get; }

        /// <summary>
        /// SupportByVersion Word 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff193753.aspx </remarks>
        [SupportByVersion("Word", 14, 15, 16)]
        NetOffice.WordApi.Walls BackWall { get; }

        /// <summary>
        /// SupportByVersion Word 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff195916.aspx </remarks>
        [SupportByVersion("Word", 14, 15, 16)]
        object ChartStyle { get; set; }

        /// <summary>
        /// SupportByVersion Word 14, 15, 16
        /// Get
        /// Unknown COM Proxy
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff836370.aspx </remarks>
        [SupportByVersion("Word", 14, 15, 16), ProxyResult]
        object PivotLayout { get; }

        /// <summary>
        /// SupportByVersion Word 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Word", 14, 15, 16)]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        bool HasPivotFields { get; set; }

        /// <summary>
        /// SupportByVersion Word 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff193871.aspx </remarks>
        [SupportByVersion("Word", 14, 15, 16)]
        bool ShowDataLabelsOverMaximum { get; set; }

        /// <summary>
        /// SupportByVersion Word 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838941.aspx </remarks>
        [SupportByVersion("Word", 14, 15, 16)]
        NetOffice.WordApi.ChartData ChartData { get; }

        /// <summary>
        /// SupportByVersion Word 14, 15, 16
        /// Get
        /// Unknown COM Proxy
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff837462.aspx </remarks>
        [SupportByVersion("Word", 14, 15, 16), ProxyResult]
        object Shapes { get; }

        /// <summary>
        /// SupportByVersion Word 14, 15, 16
        /// Get
        /// Unknown COM Proxy
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff835828.aspx </remarks>
        [SupportByVersion("Word", 14, 15, 16), ProxyResult]
        object Application { get; }

        /// <summary>
        /// SupportByVersion Word 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff192382.aspx </remarks>
        [SupportByVersion("Word", 14, 15, 16)]
        Int32 Creator { get; }

        /// <summary>
        /// SupportByVersion Word 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Word", 14, 15, 16)]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        NetOffice.WordApi.ChartGroup Area3DGroup { get; }

        /// <summary>
        /// SupportByVersion Word 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Word", 14, 15, 16)]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        NetOffice.WordApi.ChartGroup Bar3DGroup { get; }

        /// <summary>
        /// SupportByVersion Word 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Word", 14, 15, 16)]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        NetOffice.WordApi.ChartGroup Column3DGroup { get; }

        /// <summary>
        /// SupportByVersion Word 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Word", 14, 15, 16)]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        NetOffice.WordApi.ChartGroup Line3DGroup { get; }

        /// <summary>
        /// SupportByVersion Word 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Word", 14, 15, 16)]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        NetOffice.WordApi.ChartGroup Pie3DGroup { get; }

        /// <summary>
        /// SupportByVersion Word 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Word", 14, 15, 16)]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        NetOffice.WordApi.ChartGroup SurfaceGroup { get; }

        /// <summary>
        /// SupportByVersion Word 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff198336.aspx </remarks>
        [SupportByVersion("Word", 14, 15, 16)]
        bool ShowReportFilterFieldButtons { get; set; }

        /// <summary>
        /// SupportByVersion Word 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff839101.aspx </remarks>
        [SupportByVersion("Word", 14, 15, 16)]
        bool ShowLegendFieldButtons { get; set; }

        /// <summary>
        /// SupportByVersion Word 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff845234.aspx </remarks>
        [SupportByVersion("Word", 14, 15, 16)]
        bool ShowAxisFieldButtons { get; set; }

        /// <summary>
        /// SupportByVersion Word 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff834948.aspx </remarks>
        [SupportByVersion("Word", 14, 15, 16)]
        bool ShowValueFieldButtons { get; set; }

        /// <summary>
        /// SupportByVersion Word 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff834564.aspx </remarks>
        [SupportByVersion("Word", 14, 15, 16)]
        bool ShowAllFieldButtons { get; set; }

        /// <summary>
        /// SupportByVersion Word 15,16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/jj230485.aspx </remarks>
        [SupportByVersion("Word", 15, 16)]
        NetOffice.WordApi.Enums.XlCategoryLabelLevel CategoryLabelLevel { get; set; }

        /// <summary>
        /// SupportByVersion Word 15,16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/jj232218.aspx </remarks>
        [SupportByVersion("Word", 15, 16)]
        NetOffice.WordApi.Enums.XlSeriesNameLevel SeriesNameLevel { get; set; }

        /// <summary>
        /// SupportByVersion Word 15,16
        /// Get
        /// </summary>
        [SupportByVersion("Word", 15, 16)]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        bool HasHiddenContent { get; }

        /// <summary>
        /// SupportByVersion Word 15,16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/jj231924.aspx </remarks>
        [SupportByVersion("Word", 15, 16)]
        object ChartColor { get; set; }

        #endregion

        #region Methods

        /// <summary>
        /// SupportByVersion Word 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff837270.aspx </remarks>
        /// <param name="index">optional object index</param>
        [SupportByVersion("Word", 14, 15, 16)]
        object SeriesCollection(object index);

        /// <summary>
        /// SupportByVersion Word 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff837270.aspx </remarks>
        [CustomMethod]
        [SupportByVersion("Word", 14, 15, 16)]
        object SeriesCollection();

        /// <summary>
        /// SupportByVersion Word 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194055.aspx </remarks>
        /// <param name="type">optional NetOffice.WordApi.Enums.XlDataLabelsType Type = 2</param>
        /// <param name="legendKey">optional object legendKey</param>
        /// <param name="autoText">optional object autoText</param>
        /// <param name="hasLeaderLines">optional object hasLeaderLines</param>
        /// <param name="showSeriesName">optional object showSeriesName</param>
        /// <param name="showCategoryName">optional object showCategoryName</param>
        /// <param name="showValue">optional object showValue</param>
        /// <param name="showPercentage">optional object showPercentage</param>
        /// <param name="showBubbleSize">optional object showBubbleSize</param>
        /// <param name="separator">optional object separator</param>
        [SupportByVersion("Word", 14, 15, 16)]
        void ApplyDataLabels(object type, object legendKey, object autoText, object hasLeaderLines, object showSeriesName, object showCategoryName, object showValue, object showPercentage, object showBubbleSize, object separator);

        /// <summary>
        /// SupportByVersion Word 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194055.aspx </remarks>
        [CustomMethod]
        [SupportByVersion("Word", 14, 15, 16)]
        void ApplyDataLabels();

        /// <summary>
        /// SupportByVersion Word 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194055.aspx </remarks>
        /// <param name="type">optional NetOffice.WordApi.Enums.XlDataLabelsType Type = 2</param>
        [CustomMethod]
        [SupportByVersion("Word", 14, 15, 16)]
        void ApplyDataLabels(object type);

        /// <summary>
        /// SupportByVersion Word 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194055.aspx </remarks>
        /// <param name="type">optional NetOffice.WordApi.Enums.XlDataLabelsType Type = 2</param>
        /// <param name="legendKey">optional object legendKey</param>
        [CustomMethod]
        [SupportByVersion("Word", 14, 15, 16)]
        void ApplyDataLabels(object type, object legendKey);

        /// <summary>
        /// SupportByVersion Word 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194055.aspx </remarks>
        /// <param name="type">optional NetOffice.WordApi.Enums.XlDataLabelsType Type = 2</param>
        /// <param name="legendKey">optional object legendKey</param>
        /// <param name="autoText">optional object autoText</param>
        [CustomMethod]
        [SupportByVersion("Word", 14, 15, 16)]
        void ApplyDataLabels(object type, object legendKey, object autoText);

        /// <summary>
        /// SupportByVersion Word 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194055.aspx </remarks>
        /// <param name="type">optional NetOffice.WordApi.Enums.XlDataLabelsType Type = 2</param>
        /// <param name="legendKey">optional object legendKey</param>
        /// <param name="autoText">optional object autoText</param>
        /// <param name="hasLeaderLines">optional object hasLeaderLines</param>
        [CustomMethod]
        [SupportByVersion("Word", 14, 15, 16)]
        void ApplyDataLabels(object type, object legendKey, object autoText, object hasLeaderLines);

        /// <summary>
        /// SupportByVersion Word 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194055.aspx </remarks>
        /// <param name="type">optional NetOffice.WordApi.Enums.XlDataLabelsType Type = 2</param>
        /// <param name="legendKey">optional object legendKey</param>
        /// <param name="autoText">optional object autoText</param>
        /// <param name="hasLeaderLines">optional object hasLeaderLines</param>
        /// <param name="showSeriesName">optional object showSeriesName</param>
        [CustomMethod]
        [SupportByVersion("Word", 14, 15, 16)]
        void ApplyDataLabels(object type, object legendKey, object autoText, object hasLeaderLines, object showSeriesName);

        /// <summary>
        /// SupportByVersion Word 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194055.aspx </remarks>
        /// <param name="type">optional NetOffice.WordApi.Enums.XlDataLabelsType Type = 2</param>
        /// <param name="legendKey">optional object legendKey</param>
        /// <param name="autoText">optional object autoText</param>
        /// <param name="hasLeaderLines">optional object hasLeaderLines</param>
        /// <param name="showSeriesName">optional object showSeriesName</param>
        /// <param name="showCategoryName">optional object showCategoryName</param>
        [CustomMethod]
        [SupportByVersion("Word", 14, 15, 16)]
        void ApplyDataLabels(object type, object legendKey, object autoText, object hasLeaderLines, object showSeriesName, object showCategoryName);

        /// <summary>
        /// SupportByVersion Word 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194055.aspx </remarks>
        /// <param name="type">optional NetOffice.WordApi.Enums.XlDataLabelsType Type = 2</param>
        /// <param name="legendKey">optional object legendKey</param>
        /// <param name="autoText">optional object autoText</param>
        /// <param name="hasLeaderLines">optional object hasLeaderLines</param>
        /// <param name="showSeriesName">optional object showSeriesName</param>
        /// <param name="showCategoryName">optional object showCategoryName</param>
        /// <param name="showValue">optional object showValue</param>
        [CustomMethod]
        [SupportByVersion("Word", 14, 15, 16)]
        void ApplyDataLabels(object type, object legendKey, object autoText, object hasLeaderLines, object showSeriesName, object showCategoryName, object showValue);

        /// <summary>
        /// SupportByVersion Word 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194055.aspx </remarks>
        /// <param name="type">optional NetOffice.WordApi.Enums.XlDataLabelsType Type = 2</param>
        /// <param name="legendKey">optional object legendKey</param>
        /// <param name="autoText">optional object autoText</param>
        /// <param name="hasLeaderLines">optional object hasLeaderLines</param>
        /// <param name="showSeriesName">optional object showSeriesName</param>
        /// <param name="showCategoryName">optional object showCategoryName</param>
        /// <param name="showValue">optional object showValue</param>
        /// <param name="showPercentage">optional object showPercentage</param>
        [CustomMethod]
        [SupportByVersion("Word", 14, 15, 16)]
        void ApplyDataLabels(object type, object legendKey, object autoText, object hasLeaderLines, object showSeriesName, object showCategoryName, object showValue, object showPercentage);

        /// <summary>
        /// SupportByVersion Word 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194055.aspx </remarks>
        /// <param name="type">optional NetOffice.WordApi.Enums.XlDataLabelsType Type = 2</param>
        /// <param name="legendKey">optional object legendKey</param>
        /// <param name="autoText">optional object autoText</param>
        /// <param name="hasLeaderLines">optional object hasLeaderLines</param>
        /// <param name="showSeriesName">optional object showSeriesName</param>
        /// <param name="showCategoryName">optional object showCategoryName</param>
        /// <param name="showValue">optional object showValue</param>
        /// <param name="showPercentage">optional object showPercentage</param>
        /// <param name="showBubbleSize">optional object showBubbleSize</param>
        [CustomMethod]
        [SupportByVersion("Word", 14, 15, 16)]
        void ApplyDataLabels(object type, object legendKey, object autoText, object hasLeaderLines, object showSeriesName, object showCategoryName, object showValue, object showPercentage, object showBubbleSize);

        /// <summary>
        /// SupportByVersion Word 14, 15, 16
        /// </summary>
        /// <param name="chartType">NetOffice.OfficeApi.Enums.XlChartType chartType</param>
        /// <param name="typeName">optional object typeName</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [SupportByVersion("Word", 14, 15, 16)]
        void ApplyCustomType(NetOffice.OfficeApi.Enums.XlChartType chartType, object typeName);

        /// <summary>
        /// SupportByVersion Word 14, 15, 16
        /// </summary>
        /// <param name="chartType">NetOffice.OfficeApi.Enums.XlChartType chartType</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [CustomMethod]
        [SupportByVersion("Word", 14, 15, 16)]
        void ApplyCustomType(NetOffice.OfficeApi.Enums.XlChartType chartType);

        /// <summary>
        /// SupportByVersion Word 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff839889.aspx </remarks>
        /// <param name="x">Int32 x</param>
        /// <param name="y">Int32 y</param>
        /// <param name="elementID">Int32 elementID</param>
        /// <param name="arg1">Int32 arg1</param>
        /// <param name="arg2">Int32 arg2</param>
        [SupportByVersion("Word", 14, 15, 16)]
        void GetChartElement(Int32 x, Int32 y, out Int32 elementID, out Int32 arg1, out Int32 arg2);

        /// <summary>
        /// SupportByVersion Word 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff822921.aspx </remarks>
        /// <param name="source">string source</param>
        /// <param name="plotBy">optional object plotBy</param>
        [SupportByVersion("Word", 14, 15, 16)]
        void SetSourceData(string source, object plotBy);

        /// <summary>
        /// SupportByVersion Word 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff822921.aspx </remarks>
        /// <param name="source">string source</param>
        [CustomMethod]
        [SupportByVersion("Word", 14, 15, 16)]
        void SetSourceData(string source);

        /// <summary>
        /// SupportByVersion Word 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff193697.aspx </remarks>
        /// <param name="type">optional object type</param>
        /// <param name="axisGroup">optional NetOffice.WordApi.Enums.XlAxisGroup AxisGroup = 1</param>
        [SupportByVersion("Word", 14, 15, 16)]
        object Axes(object type, object axisGroup);

        /// <summary>
        /// SupportByVersion Word 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff193697.aspx </remarks>
        [CustomMethod]
        [SupportByVersion("Word", 14, 15, 16)]
        object Axes();

        /// <summary>
        /// SupportByVersion Word 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff193697.aspx </remarks>
        /// <param name="type">optional object type</param>
        [CustomMethod]
        [SupportByVersion("Word", 14, 15, 16)]
        object Axes(object type);

        /// <summary>
        /// SupportByVersion Word 14, 15, 16
        /// </summary>
        /// <param name="gallery">Int32 gallery</param>
        /// <param name="format">optional object format</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [SupportByVersion("Word", 14, 15, 16)]
        void AutoFormat(Int32 gallery, object format);

        /// <summary>
        /// SupportByVersion Word 14, 15, 16
        /// </summary>
        /// <param name="gallery">Int32 gallery</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [CustomMethod]
        [SupportByVersion("Word", 14, 15, 16)]
        void AutoFormat(Int32 gallery);

        /// <summary>
        /// SupportByVersion Word 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff197864.aspx </remarks>
        /// <param name="fileName">string fileName</param>
        [SupportByVersion("Word", 14, 15, 16)]
        void SetBackgroundPicture(string fileName);

        /// <summary>
        /// SupportByVersion Word 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff196615.aspx </remarks>
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
        [SupportByVersion("Word", 14, 15, 16)]
        void ChartWizard(object source, object gallery, object format, object plotBy, object categoryLabels, object seriesLabels, object hasLegend, object title, object categoryTitle, object valueTitle, object extraTitle);

        /// <summary>
        /// SupportByVersion Word 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff196615.aspx </remarks>
        [CustomMethod]
        [SupportByVersion("Word", 14, 15, 16)]
        void ChartWizard();

        /// <summary>
        /// SupportByVersion Word 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff196615.aspx </remarks>
        /// <param name="source">optional object source</param>
        [CustomMethod]
        [SupportByVersion("Word", 14, 15, 16)]
        void ChartWizard(object source);

        /// <summary>
        /// SupportByVersion Word 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff196615.aspx </remarks>
        /// <param name="source">optional object source</param>
        /// <param name="gallery">optional object gallery</param>
        [CustomMethod]
        [SupportByVersion("Word", 14, 15, 16)]
        void ChartWizard(object source, object gallery);

        /// <summary>
        /// SupportByVersion Word 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff196615.aspx </remarks>
        /// <param name="source">optional object source</param>
        /// <param name="gallery">optional object gallery</param>
        /// <param name="format">optional object format</param>
        [CustomMethod]
        [SupportByVersion("Word", 14, 15, 16)]
        void ChartWizard(object source, object gallery, object format);

        /// <summary>
        /// SupportByVersion Word 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff196615.aspx </remarks>
        /// <param name="source">optional object source</param>
        /// <param name="gallery">optional object gallery</param>
        /// <param name="format">optional object format</param>
        /// <param name="plotBy">optional object plotBy</param>
        [CustomMethod]
        [SupportByVersion("Word", 14, 15, 16)]
        void ChartWizard(object source, object gallery, object format, object plotBy);

        /// <summary>
        /// SupportByVersion Word 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff196615.aspx </remarks>
        /// <param name="source">optional object source</param>
        /// <param name="gallery">optional object gallery</param>
        /// <param name="format">optional object format</param>
        /// <param name="plotBy">optional object plotBy</param>
        /// <param name="categoryLabels">optional object categoryLabels</param>
        [CustomMethod]
        [SupportByVersion("Word", 14, 15, 16)]
        void ChartWizard(object source, object gallery, object format, object plotBy, object categoryLabels);

        /// <summary>
        /// SupportByVersion Word 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff196615.aspx </remarks>
        /// <param name="source">optional object source</param>
        /// <param name="gallery">optional object gallery</param>
        /// <param name="format">optional object format</param>
        /// <param name="plotBy">optional object plotBy</param>
        /// <param name="categoryLabels">optional object categoryLabels</param>
        /// <param name="seriesLabels">optional object seriesLabels</param>
        [CustomMethod]
        [SupportByVersion("Word", 14, 15, 16)]
        void ChartWizard(object source, object gallery, object format, object plotBy, object categoryLabels, object seriesLabels);

        /// <summary>
        /// SupportByVersion Word 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff196615.aspx </remarks>
        /// <param name="source">optional object source</param>
        /// <param name="gallery">optional object gallery</param>
        /// <param name="format">optional object format</param>
        /// <param name="plotBy">optional object plotBy</param>
        /// <param name="categoryLabels">optional object categoryLabels</param>
        /// <param name="seriesLabels">optional object seriesLabels</param>
        /// <param name="hasLegend">optional object hasLegend</param>
        [CustomMethod]
        [SupportByVersion("Word", 14, 15, 16)]
        void ChartWizard(object source, object gallery, object format, object plotBy, object categoryLabels, object seriesLabels, object hasLegend);

        /// <summary>
        /// SupportByVersion Word 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff196615.aspx </remarks>
        /// <param name="source">optional object source</param>
        /// <param name="gallery">optional object gallery</param>
        /// <param name="format">optional object format</param>
        /// <param name="plotBy">optional object plotBy</param>
        /// <param name="categoryLabels">optional object categoryLabels</param>
        /// <param name="seriesLabels">optional object seriesLabels</param>
        /// <param name="hasLegend">optional object hasLegend</param>
        /// <param name="title">optional object title</param>
        [CustomMethod]
        [SupportByVersion("Word", 14, 15, 16)]
        void ChartWizard(object source, object gallery, object format, object plotBy, object categoryLabels, object seriesLabels, object hasLegend, object title);

        /// <summary>
        /// SupportByVersion Word 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff196615.aspx </remarks>
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
        [SupportByVersion("Word", 14, 15, 16)]
        void ChartWizard(object source, object gallery, object format, object plotBy, object categoryLabels, object seriesLabels, object hasLegend, object title, object categoryTitle);

        /// <summary>
        /// SupportByVersion Word 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff196615.aspx </remarks>
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
        [SupportByVersion("Word", 14, 15, 16)]
        void ChartWizard(object source, object gallery, object format, object plotBy, object categoryLabels, object seriesLabels, object hasLegend, object title, object categoryTitle, object valueTitle);

        /// <summary>
        /// SupportByVersion Word 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff823258.aspx </remarks>
        /// <param name="appearance">optional NetOffice.WordApi.Enums.XlPictureAppearance Appearance = 1</param>
        /// <param name="format">optional NetOffice.WordApi.Enums.XlCopyPictureFormat Format = -4147</param>
        /// <param name="size">optional NetOffice.WordApi.Enums.XlPictureAppearance Size = 2</param>
        [SupportByVersion("Word", 14, 15, 16)]
        void CopyPicture(object appearance, object format, object size);

        /// <summary>
        /// SupportByVersion Word 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff823258.aspx </remarks>
        [CustomMethod]
        [SupportByVersion("Word", 14, 15, 16)]
        void CopyPicture();

        /// <summary>
        /// SupportByVersion Word 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff823258.aspx </remarks>
        /// <param name="appearance">optional NetOffice.WordApi.Enums.XlPictureAppearance Appearance = 1</param>
        [CustomMethod]
        [SupportByVersion("Word", 14, 15, 16)]
        void CopyPicture(object appearance);

        /// <summary>
        /// SupportByVersion Word 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff823258.aspx </remarks>
        /// <param name="appearance">optional NetOffice.WordApi.Enums.XlPictureAppearance Appearance = 1</param>
        /// <param name="format">optional NetOffice.WordApi.Enums.XlCopyPictureFormat Format = -4147</param>
        [CustomMethod]
        [SupportByVersion("Word", 14, 15, 16)]
        void CopyPicture(object appearance, object format);

        /// <summary>
        /// SupportByVersion Word 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff839410.aspx </remarks>
        /// <param name="type">optional object type</param>
        [SupportByVersion("Word", 14, 15, 16)]
        void Paste(object type);

        /// <summary>
        /// SupportByVersion Word 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff839410.aspx </remarks>
        [CustomMethod]
        [SupportByVersion("Word", 14, 15, 16)]
        void Paste();

        /// <summary>
        /// SupportByVersion Word 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff195106.aspx </remarks>
        /// <param name="fileName">string fileName</param>
        /// <param name="filterName">optional object filterName</param>
        /// <param name="interactive">optional object interactive</param>
        [SupportByVersion("Word", 14, 15, 16)]
        bool Export(string fileName, object filterName, object interactive);

        /// <summary>
        /// SupportByVersion Word 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff195106.aspx </remarks>
        /// <param name="fileName">string fileName</param>
        [CustomMethod]
        [SupportByVersion("Word", 14, 15, 16)]
        bool Export(string fileName);

        /// <summary>
        /// SupportByVersion Word 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff195106.aspx </remarks>
        /// <param name="fileName">string fileName</param>
        /// <param name="filterName">optional object filterName</param>
        [CustomMethod]
        [SupportByVersion("Word", 14, 15, 16)]
        bool Export(string fileName, object filterName);

        /// <summary>
        /// SupportByVersion Word 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff839841.aspx </remarks>
        /// <param name="name">object name</param>
        [SupportByVersion("Word", 14, 15, 16)]
        void SetDefaultChart(object name);

        /// <summary>
        /// SupportByVersion Word 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff845631.aspx </remarks>
        /// <param name="fileName">string fileName</param>
        [SupportByVersion("Word", 14, 15, 16)]
        void ApplyChartTemplate(string fileName);

        /// <summary>
        /// SupportByVersion Word 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff839083.aspx </remarks>
        /// <param name="fileName">string fileName</param>
        [SupportByVersion("Word", 14, 15, 16)]
        void SaveChartTemplate(string fileName);

        /// <summary>
        /// SupportByVersion Word 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff193390.aspx </remarks>
        [SupportByVersion("Word", 14, 15, 16)]
        void ClearToMatchStyle();

        /// <summary>
        /// SupportByVersion Word 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff840397.aspx </remarks>
        /// <param name="layout">Int32 layout</param>
        /// <param name="chartType">optional object chartType</param>
        [SupportByVersion("Word", 14, 15, 16)]
        void ApplyLayout(Int32 layout, object chartType);

        /// <summary>
        /// SupportByVersion Word 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff840397.aspx </remarks>
        /// <param name="layout">Int32 layout</param>
        [CustomMethod]
        [SupportByVersion("Word", 14, 15, 16)]
        void ApplyLayout(Int32 layout);

        /// <summary>
        /// SupportByVersion Word 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff192135.aspx </remarks>
        [SupportByVersion("Word", 14, 15, 16)]
        void Refresh();

        /// <summary>
        /// SupportByVersion Word 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838552.aspx </remarks>
        /// <param name="element">NetOffice.OfficeApi.Enums.MsoChartElementType element</param>
        [SupportByVersion("Word", 14, 15, 16)]
        void SetElement(NetOffice.OfficeApi.Enums.MsoChartElementType element);

        /// <summary>
        /// SupportByVersion Word 14, 15, 16
        /// </summary>
        /// <param name="index">optional object index</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [SupportByVersion("Word", 14, 15, 16)]
        object AreaGroups(object index);

        /// <summary>
        /// SupportByVersion Word 14, 15, 16
        /// </summary>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [CustomMethod]
        [SupportByVersion("Word", 14, 15, 16)]
        object AreaGroups();

        /// <summary>
        /// SupportByVersion Word 14, 15, 16
        /// </summary>
        /// <param name="index">optional object index</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [SupportByVersion("Word", 14, 15, 16)]
        object BarGroups(object index);

        /// <summary>
        /// SupportByVersion Word 14, 15, 16
        /// </summary>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [CustomMethod]
        [SupportByVersion("Word", 14, 15, 16)]
        object BarGroups();

        /// <summary>
        /// SupportByVersion Word 14, 15, 16
        /// </summary>
        /// <param name="index">optional object index</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [SupportByVersion("Word", 14, 15, 16)]
        object ColumnGroups(object index);

        /// <summary>
        /// SupportByVersion Word 14, 15, 16
        /// </summary>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [CustomMethod]
        [SupportByVersion("Word", 14, 15, 16)]
        object ColumnGroups();

        /// <summary>
        /// SupportByVersion Word 14, 15, 16
        /// </summary>
        /// <param name="index">optional object index</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [SupportByVersion("Word", 14, 15, 16)]
        object LineGroups(object index);

        /// <summary>
        /// SupportByVersion Word 14, 15, 16
        /// </summary>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [CustomMethod]
        [SupportByVersion("Word", 14, 15, 16)]
        object LineGroups();

        /// <summary>
        /// SupportByVersion Word 14, 15, 16
        /// </summary>
        /// <param name="index">optional object index</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [SupportByVersion("Word", 14, 15, 16)]
        object PieGroups(object index);

        /// <summary>
        /// SupportByVersion Word 14, 15, 16
        /// </summary>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [CustomMethod]
        [SupportByVersion("Word", 14, 15, 16)]
        object PieGroups();

        /// <summary>
        /// SupportByVersion Word 14, 15, 16
        /// </summary>
        /// <param name="index">optional object index</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [SupportByVersion("Word", 14, 15, 16)]
        object DoughnutGroups(object index);

        /// <summary>
        /// SupportByVersion Word 14, 15, 16
        /// </summary>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [CustomMethod]
        [SupportByVersion("Word", 14, 15, 16)]
        object DoughnutGroups();

        /// <summary>
        /// SupportByVersion Word 14, 15, 16
        /// </summary>
        /// <param name="index">optional object index</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [SupportByVersion("Word", 14, 15, 16)]
        object RadarGroups(object index);

        /// <summary>
        /// SupportByVersion Word 14, 15, 16
        /// </summary>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [CustomMethod]
        [SupportByVersion("Word", 14, 15, 16)]
        object RadarGroups();

        /// <summary>
        /// SupportByVersion Word 14, 15, 16
        /// </summary>
        /// <param name="index">optional object index</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [SupportByVersion("Word", 14, 15, 16)]
        object XYGroups(object index);

        /// <summary>
        /// SupportByVersion Word 14, 15, 16
        /// </summary>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [CustomMethod]
        [SupportByVersion("Word", 14, 15, 16)]
        object XYGroups();

        /// <summary>
        /// SupportByVersion Word 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff840074.aspx </remarks>
        [SupportByVersion("Word", 14, 15, 16)]
        object Delete();

        /// <summary>
        /// SupportByVersion Word 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff192354.aspx </remarks>
        /// <param name="before">optional object before</param>
        /// <param name="after">optional object after</param>
        [SupportByVersion("Word", 14, 15, 16)]
        void Copy(object before, object after);

        /// <summary>
        /// SupportByVersion Word 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff192354.aspx </remarks>
        [CustomMethod]
        [SupportByVersion("Word", 14, 15, 16)]
        void Copy();

        /// <summary>
        /// SupportByVersion Word 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff192354.aspx </remarks>
        /// <param name="before">optional object before</param>
        [CustomMethod]
        [SupportByVersion("Word", 14, 15, 16)]
        void Copy(object before);

        /// <summary>
        /// SupportByVersion Word 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff191928.aspx </remarks>
        /// <param name="replace">optional object replace</param>
        [SupportByVersion("Word", 14, 15, 16)]
        object Select(object replace);

        /// <summary>
        /// SupportByVersion Word 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff191928.aspx </remarks>
        [CustomMethod]
        [SupportByVersion("Word", 14, 15, 16)]
        object Select();

        /// <summary>
        /// SupportByVersion Word 15,16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/jj229848.aspx </remarks>
        /// <param name="index">optional object index</param>
        [SupportByVersion("Word", 15, 16)]
        object FullSeriesCollection(object index);

        /// <summary>
        /// SupportByVersion Word 15,16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/jj229848.aspx </remarks>
        [CustomMethod]
        [SupportByVersion("Word", 15, 16)]
        object FullSeriesCollection();

        /// <summary>
        /// SupportByVersion Word 15,16
        /// </summary>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [SupportByVersion("Word", 15, 16)]
        void DeleteHiddenContent();

        /// <summary>
        /// SupportByVersion Word 15,16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/jj230203.aspx </remarks>
        [SupportByVersion("Word", 15, 16)]
        void ClearToMatchColorStyle();

        #endregion
    }
}
