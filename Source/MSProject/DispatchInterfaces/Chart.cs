using System.Reflection;
using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.MSProjectApi
{
    /// <summary>
    /// Chart
    /// </summary>
    [SyntaxBypass]
    public interface Chart_ : ICOMObject
    {
        #region Properties

        /// <summary>
        /// SupportByVersion MSProject 11
        /// Get
        /// Unknown COM Proxy
        /// </summary>
        /// <param name="pvarIndex">optional object pvarIndex</param>
        /// <param name="varIgallery">optional object varIgallery</param>
        [SupportByVersion("MSProject", 11), ProxyResult]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        object get_ChartGroups(object pvarIndex, object varIgallery);

        /// <summary>
        /// SupportByVersion MSProject 11
        /// Alias for get_ChartGroups
        /// Unknown COM Proxy
        /// </summary>
        /// <param name="pvarIndex">optional object pvarIndex</param>
        /// <param name="varIgallery">optional object varIgallery</param>
        [SupportByVersion("MSProject", 11), ProxyResult, Redirect("get_ChartGroups")]
        object ChartGroups(object pvarIndex, object varIgallery);

        /// <summary>
        /// SupportByVersion MSProject 11
        /// Get
        /// Unknown COM Proxy
        /// </summary>
        /// <param name="pvarIndex">optional object pvarIndex</param>
        [SupportByVersion("MSProject", 11), ProxyResult]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        object get_ChartGroups(object pvarIndex);

        /// <summary>
        /// SupportByVersion MSProject 11
        /// Alias for get_ChartGroups
        /// Unknown COM Proxy
        /// </summary>
        /// <param name="pvarIndex">optional object pvarIndex</param>
        [SupportByVersion("MSProject", 11), ProxyResult, Redirect("get_ChartGroups")]
        object ChartGroups(object pvarIndex);

        /// <summary>
        /// SupportByVersion MSProject 11
        /// Get/Set
        /// </summary>
        /// <param name="axisType">optional object axisType</param>
        /// <param name="axisGroup">optional object axisGroup</param>
        [SupportByVersion("MSProject", 11)]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        object get_HasAxis(object axisType, object axisGroup);

        /// <summary>
        /// SupportByVersion MSProject 11
        /// Get/Set
        /// </summary>
        /// <param name="axisType">optional object axisType</param>
        /// <param name="axisGroup">optional object axisGroup</param>
        /// <param name="value">optional object value</param>
        [SupportByVersion("MSProject", 11)]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        void set_HasAxis(object axisType, object axisGroup, object value);

        /// <summary>
        /// SupportByVersion MSProject 11
        /// Alias for get_HasAxis
        /// </summary>
        /// <param name="axisType">optional object axisType</param>
        /// <param name="axisGroup">optional object axisGroup</param>
        [SupportByVersion("MSProject", 11), Redirect("get_HasAxis")]
        object HasAxis(object axisType, object axisGroup);

        /// <summary>
        /// SupportByVersion MSProject 11
        /// Get/Set
        /// </summary>
        /// <param name="axisType">optional object axisType</param>
        [SupportByVersion("MSProject", 11)]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        object get_HasAxis(object axisType);

        /// <summary>
        /// SupportByVersion MSProject 11
        /// Get/Set
        /// </summary>
        /// <param name="axisType">optional object axisType</param>
        /// <param name="value">optional object value</param>
        [SupportByVersion("MSProject", 11)]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        void set_HasAxis(object axisType, object value);

        /// <summary>
        /// SupportByVersion MSProject 11
        /// Alias for get_HasAxis
        /// </summary>
        /// <param name="axisType">optional object axisType</param>
        [SupportByVersion("MSProject", 11), Redirect("get_HasAxis")]
        object HasAxis(object axisType);

        /// <summary>
        /// SupportByVersion MSProject 11
        /// Get
        /// </summary>
        /// <param name="fBackWall">optional bool fBackWall</param>
        [SupportByVersion("MSProject", 11)]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        NetOffice.OfficeApi.IMsoWalls get_Walls(object fBackWall);

        /// <summary>
        /// SupportByVersion MSProject 11
        /// Alias for get_Walls
        /// </summary>
        /// <param name="fBackWall">optional bool fBackWall</param>
        [SupportByVersion("MSProject", 11), Redirect("get_Walls")]
        NetOffice.OfficeApi.IMsoWalls Walls(object fBackWall);

        #endregion
    }
   
    /// <summary>
    /// DispatchInterface Chart 
    /// SupportByVersion MSProject, 11
    /// </summary>
    [SupportByVersion("MSProject", 11)]
    [EntityType(EntityType.IsDispatchInterface)]
	[TypeId("CA6893A3-E8B7-46ED-89AB-0600354CBD7B")]
    public interface Chart : Chart_
    {
        #region Properties

        /// <summary>
        /// SupportByVersion MSProject 11
        /// Get
        /// Unknown COM Proxy
        /// </summary>
        [SupportByVersion("MSProject", 11), ProxyResult]
        object Parent { get; }

        /// <summary>
        /// SupportByVersion MSProject 11
        /// Get/Set
        /// </summary>
        [SupportByVersion("MSProject", 11)]
        bool HasTitle { get; set; }

        /// <summary>
        /// SupportByVersion MSProject 11
        /// Get
        /// </summary>
        [SupportByVersion("MSProject", 11)]
        [BaseResult]
        NetOffice.OfficeApi.IMsoChartTitle ChartTitle { get; }

        /// <summary>
        /// SupportByVersion MSProject 11
        /// Get/Set
        /// </summary>
        [SupportByVersion("MSProject", 11)]
        Int32 DepthPercent { get; set; }

        /// <summary>
        /// SupportByVersion MSProject 11
        /// Get/Set
        /// </summary>
        [SupportByVersion("MSProject", 11)]
        Int32 Elevation { get; set; }

        /// <summary>
        /// SupportByVersion MSProject 11
        /// Get/Set
        /// </summary>
        [SupportByVersion("MSProject", 11)]
        Int32 GapDepth { get; set; }

        /// <summary>
        /// SupportByVersion MSProject 11
        /// Get/Set
        /// </summary>
        [SupportByVersion("MSProject", 11)]
        Int32 HeightPercent { get; set; }

        /// <summary>
        /// SupportByVersion MSProject 11
        /// Get/Set
        /// </summary>
        [SupportByVersion("MSProject", 11)]
        Int32 Perspective { get; set; }

        /// <summary>
        /// SupportByVersion MSProject 11
        /// Get/Set
        /// </summary>
        [SupportByVersion("MSProject", 11)]
        object RightAngleAxes { get; set; }

        /// <summary>
        /// SupportByVersion MSProject 11
        /// Get/Set
        /// </summary>
        [SupportByVersion("MSProject", 11)]
        object Rotation { get; set; }

        /// <summary>
        /// SupportByVersion MSProject 11
        /// Get/Set
        /// </summary>
        [SupportByVersion("MSProject", 11)]
        NetOffice.OfficeApi.Enums.XlDisplayBlanksAs DisplayBlanksAs { get; set; }

        /// <summary>
        /// SupportByVersion MSProject 11
        /// Get/Set
        /// </summary>
        [SupportByVersion("MSProject", 11)]
        bool ProtectData { get; set; }

        /// <summary>
        /// SupportByVersion MSProject 11
        /// Get/Set
        /// </summary>
        [SupportByVersion("MSProject", 11)]
        bool ProtectFormatting { get; set; }

        /// <summary>
        /// SupportByVersion MSProject 11
        /// Get/Set
        /// </summary>
        [SupportByVersion("MSProject", 11)]
        bool ProtectGoalSeek { get; set; }

        /// <summary>
        /// SupportByVersion MSProject 11
        /// Get/Set
        /// </summary>
        [SupportByVersion("MSProject", 11)]
        bool ProtectSelection { get; set; }

        /// <summary>
        /// SupportByVersion MSProject 11
        /// Get/Set
        /// </summary>
        [SupportByVersion("MSProject", 11)]
        bool ProtectChartObjects { get; set; }

        /// <summary>
        /// SupportByVersion MSProject 11
        /// Get
        /// Unknown COM Proxy
        /// </summary>
        [SupportByVersion("MSProject", 11), ProxyResult]
        new object ChartGroups { get; }

        /// <summary>
        /// SupportByVersion MSProject 11
        /// Get/Set
        /// </summary>
        [SupportByVersion("MSProject", 11)]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        Int32 SubType { get; set; }

        /// <summary>
        /// SupportByVersion MSProject 11
        /// Get/Set
        /// </summary>
        [SupportByVersion("MSProject", 11)]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        Int32 Type { get; set; }

        /// <summary>
        /// SupportByVersion MSProject 11
        /// Get
        /// </summary>
        [SupportByVersion("MSProject", 11)]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        NetOffice.OfficeApi.IMsoCorners Corners { get; }

        /// <summary>
        /// SupportByVersion MSProject 11
        /// Get/Set
        /// </summary>
        [SupportByVersion("MSProject", 11)]
        NetOffice.OfficeApi.Enums.XlChartType ChartType { get; set; }

        /// <summary>
        /// SupportByVersion MSProject 11
        /// Get/Set
        /// </summary>
        [SupportByVersion("MSProject", 11)]
        bool HasDataTable { get; set; }

        /// <summary>
        /// SupportByVersion MSProject 11
        /// Get/Set
        /// </summary>
        [SupportByVersion("MSProject", 11)]
        NetOffice.OfficeApi.Enums.XlRowCol PlotBy { get; set; }

        /// <summary>
        /// SupportByVersion MSProject 11
        /// Get/Set
        /// </summary>
        [SupportByVersion("MSProject", 11)]
        bool HasLegend { get; set; }

        /// <summary>
        /// SupportByVersion MSProject 11
        /// Get
        /// </summary>
        [SupportByVersion("MSProject", 11)]
        NetOffice.OfficeApi.IMsoLegend Legend { get; }

        /// <summary>
        /// SupportByVersion MSProject 11
        /// Get/Set
        /// </summary>
        [SupportByVersion("MSProject", 11)]
        new object HasAxis { get; set; }

        /// <summary>
        /// SupportByVersion MSProject 11
        /// Get
        /// </summary>
        [SupportByVersion("MSProject", 11)]
        new NetOffice.OfficeApi.IMsoWalls Walls { get; }

        /// <summary>
        /// SupportByVersion MSProject 11
        /// Get
        /// </summary>
        [SupportByVersion("MSProject", 11)]
        NetOffice.OfficeApi.IMsoFloor Floor { get; }

        /// <summary>
        /// SupportByVersion MSProject 11
        /// Get
        /// </summary>
        [SupportByVersion("MSProject", 11)]
        NetOffice.OfficeApi.IMsoPlotArea PlotArea { get; }

        /// <summary>
        /// SupportByVersion MSProject 11
        /// Get/Set
        /// </summary>
        [SupportByVersion("MSProject", 11)]
        bool PlotVisibleOnly { get; set; }

        /// <summary>
        /// SupportByVersion MSProject 11
        /// Get
        /// </summary>
        [SupportByVersion("MSProject", 11)]
        NetOffice.OfficeApi.IMsoChartArea ChartArea { get; }

        /// <summary>
        /// SupportByVersion MSProject 11
        /// Get/Set
        /// </summary>
        [SupportByVersion("MSProject", 11)]
        bool AutoScaling { get; set; }

        /// <summary>
        /// SupportByVersion MSProject 11
        /// Get
        /// </summary>
        [SupportByVersion("MSProject", 11)]
        NetOffice.OfficeApi.IMsoDataTable DataTable { get; }

        /// <summary>
        /// SupportByVersion MSProject 11
        /// Get/Set
        /// </summary>
        [SupportByVersion("MSProject", 11)]
        NetOffice.OfficeApi.Enums.XlBarShape BarShape { get; set; }

        /// <summary>
        /// SupportByVersion MSProject 11
        /// Get
        /// </summary>
        [SupportByVersion("MSProject", 11)]
        NetOffice.OfficeApi.IMsoWalls SideWall { get; }

        /// <summary>
        /// SupportByVersion MSProject 11
        /// Get
        /// </summary>
        [SupportByVersion("MSProject", 11)]
        NetOffice.OfficeApi.IMsoWalls BackWall { get; }

        /// <summary>
        /// SupportByVersion MSProject 11
        /// Get/Set
        /// </summary>
        [SupportByVersion("MSProject", 11)]
        object ChartStyle { get; set; }

        /// <summary>
        /// SupportByVersion MSProject 11
        /// Get
        /// Unknown COM Proxy
        /// </summary>
        [SupportByVersion("MSProject", 11), ProxyResult]
        object PivotLayout { get; }

        /// <summary>
        /// SupportByVersion MSProject 11
        /// Get/Set
        /// </summary>
        [SupportByVersion("MSProject", 11)]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        bool HasPivotFields { get; set; }

        /// <summary>
        /// SupportByVersion MSProject 11
        /// Get/Set
        /// </summary>
        [SupportByVersion("MSProject", 11)]
        bool ShowDataLabelsOverMaximum { get; set; }

        /// <summary>
        /// SupportByVersion MSProject 11
        /// Get
        /// Unknown COM Proxy
        /// </summary>
        [SupportByVersion("MSProject", 11), ProxyResult]
        object Selection { get; }

        /// <summary>
        /// SupportByVersion MSProject 11
        /// Get
        /// </summary>
        [SupportByVersion("MSProject", 11)]
        NetOffice.OfficeApi.IMsoChartData ChartData { get; }

        /// <summary>
        /// SupportByVersion MSProject 11
        /// Get
        /// </summary>
        [SupportByVersion("MSProject", 11)]
        NetOffice.OfficeApi.IMsoChartFormat Format { get; }

        /// <summary>
        /// SupportByVersion MSProject 11
        /// Get
        /// </summary>
        [SupportByVersion("MSProject", 11)]
        NetOffice.MSProjectApi.Shapes Shapes { get; }

        /// <summary>
        /// SupportByVersion MSProject 11
        /// Get
        /// Unknown COM Proxy
        /// </summary>
        [SupportByVersion("MSProject", 11), ProxyResult]
        object Application { get; }

        /// <summary>
        /// SupportByVersion MSProject 11
        /// Get
        /// </summary>
        [SupportByVersion("MSProject", 11)]
        Int32 Creator { get; }

        /// <summary>
        /// SupportByVersion MSProject 11
        /// Get
        /// </summary>
        [SupportByVersion("MSProject", 11)]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        NetOffice.OfficeApi.IMsoChartGroup Area3DGroup { get; }

        /// <summary>
        /// SupportByVersion MSProject 11
        /// Get
        /// </summary>
        [SupportByVersion("MSProject", 11)]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        NetOffice.OfficeApi.IMsoChartGroup Bar3DGroup { get; }

        /// <summary>
        /// SupportByVersion MSProject 11
        /// Get
        /// </summary>
        [SupportByVersion("MSProject", 11)]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        NetOffice.OfficeApi.IMsoChartGroup Column3DGroup { get; }

        /// <summary>
        /// SupportByVersion MSProject 11
        /// Get
        /// </summary>
        [SupportByVersion("MSProject", 11)]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        NetOffice.OfficeApi.IMsoChartGroup Line3DGroup { get; }

        /// <summary>
        /// SupportByVersion MSProject 11
        /// Get
        /// </summary>
        [SupportByVersion("MSProject", 11)]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        NetOffice.OfficeApi.IMsoChartGroup Pie3DGroup { get; }

        /// <summary>
        /// SupportByVersion MSProject 11
        /// Get
        /// </summary>
        [SupportByVersion("MSProject", 11)]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        NetOffice.OfficeApi.IMsoChartGroup SurfaceGroup { get; }

        /// <summary>
        /// SupportByVersion MSProject 11
        /// Get/Set
        /// </summary>
        [SupportByVersion("MSProject", 11)]
        bool ShowReportFilterFieldButtons { get; set; }

        /// <summary>
        /// SupportByVersion MSProject 11
        /// Get/Set
        /// </summary>
        [SupportByVersion("MSProject", 11)]
        bool ShowLegendFieldButtons { get; set; }

        /// <summary>
        /// SupportByVersion MSProject 11
        /// Get/Set
        /// </summary>
        [SupportByVersion("MSProject", 11)]
        bool ShowAxisFieldButtons { get; set; }

        /// <summary>
        /// SupportByVersion MSProject 11
        /// Get/Set
        /// </summary>
        [SupportByVersion("MSProject", 11)]
        bool ShowValueFieldButtons { get; set; }

        /// <summary>
        /// SupportByVersion MSProject 11
        /// Get/Set
        /// </summary>
        [SupportByVersion("MSProject", 11)]
        bool ShowAllFieldButtons { get; set; }

        /// <summary>
        /// SupportByVersion MSProject 11
        /// Get/Set
        /// </summary>
        [SupportByVersion("MSProject", 11)]
        object ChartColor { get; set; }

        #endregion

        #region Methods

        /// <summary>
        /// SupportByVersion MSProject 11
        /// </summary>
        /// <param name="password">optional object password</param>
        [SupportByVersion("MSProject", 11)]
        void UnProtect(object password);

        /// <summary>
        /// SupportByVersion MSProject 11
        /// </summary>
        [CustomMethod]
        [SupportByVersion("MSProject", 11)]
        void UnProtect();

        /// <summary>
        /// SupportByVersion MSProject 11
        /// </summary>
        /// <param name="password">optional object password</param>
        /// <param name="drawingObjects">optional object drawingObjects</param>
        /// <param name="contents">optional object contents</param>
        /// <param name="scenarios">optional object scenarios</param>
        /// <param name="userInterfaceOnly">optional object userInterfaceOnly</param>
        [SupportByVersion("MSProject", 11)]
        void Protect(object password, object drawingObjects, object contents, object scenarios, object userInterfaceOnly);

        /// <summary>
        /// SupportByVersion MSProject 11
        /// </summary>
        [CustomMethod]
        [SupportByVersion("MSProject", 11)]
        void Protect();

        /// <summary>
        /// SupportByVersion MSProject 11
        /// </summary>
        /// <param name="password">optional object password</param>
        [CustomMethod]
        [SupportByVersion("MSProject", 11)]
        void Protect(object password);

        /// <summary>
        /// SupportByVersion MSProject 11
        /// </summary>
        /// <param name="password">optional object password</param>
        /// <param name="drawingObjects">optional object drawingObjects</param>
        [CustomMethod]
        [SupportByVersion("MSProject", 11)]
        void Protect(object password, object drawingObjects);

        /// <summary>
        /// SupportByVersion MSProject 11
        /// </summary>
        /// <param name="password">optional object password</param>
        /// <param name="drawingObjects">optional object drawingObjects</param>
        /// <param name="contents">optional object contents</param>
        [CustomMethod]
        [SupportByVersion("MSProject", 11)]
        void Protect(object password, object drawingObjects, object contents);

        /// <summary>
        /// SupportByVersion MSProject 11
        /// </summary>
        /// <param name="password">optional object password</param>
        /// <param name="drawingObjects">optional object drawingObjects</param>
        /// <param name="contents">optional object contents</param>
        /// <param name="scenarios">optional object scenarios</param>
        [CustomMethod]
        [SupportByVersion("MSProject", 11)]
        void Protect(object password, object drawingObjects, object contents, object scenarios);

        /// <summary>
        /// SupportByVersion MSProject 11
        /// </summary>
        /// <param name="index">optional object index</param>
        [SupportByVersion("MSProject", 11)]
        object SeriesCollection(object index);

        /// <summary>
        /// SupportByVersion MSProject 11
        /// </summary>
        [CustomMethod]
        [SupportByVersion("MSProject", 11)]
        object SeriesCollection();

        /// <summary>
        /// SupportByVersion MSProject 11
        /// </summary>
        /// <param name="type">optional NetOffice.OfficeApi.Enums.XlDataLabelsType Type = 2</param>
        /// <param name="iMsoLegendKey">optional object iMsoLegendKey</param>
        /// <param name="autoText">optional object autoText</param>
        /// <param name="hasLeaderLines">optional object hasLeaderLines</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [SupportByVersion("MSProject", 11)]
        void _ApplyDataLabels(object type, object iMsoLegendKey, object autoText, object hasLeaderLines);

        /// <summary>
        /// SupportByVersion MSProject 11
        /// </summary>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [CustomMethod]
        [SupportByVersion("MSProject", 11)]
        void _ApplyDataLabels();

        /// <summary>
        /// SupportByVersion MSProject 11
        /// </summary>
        /// <param name="type">optional NetOffice.OfficeApi.Enums.XlDataLabelsType Type = 2</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [CustomMethod]
        [SupportByVersion("MSProject", 11)]
        void _ApplyDataLabels(object type);

        /// <summary>
        /// SupportByVersion MSProject 11
        /// </summary>
        /// <param name="type">optional NetOffice.OfficeApi.Enums.XlDataLabelsType Type = 2</param>
        /// <param name="iMsoLegendKey">optional object iMsoLegendKey</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [CustomMethod]
        [SupportByVersion("MSProject", 11)]
        void _ApplyDataLabels(object type, object iMsoLegendKey);

        /// <summary>
        /// SupportByVersion MSProject 11
        /// </summary>
        /// <param name="type">optional NetOffice.OfficeApi.Enums.XlDataLabelsType Type = 2</param>
        /// <param name="iMsoLegendKey">optional object iMsoLegendKey</param>
        /// <param name="autoText">optional object autoText</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [CustomMethod]
        [SupportByVersion("MSProject", 11)]
        void _ApplyDataLabels(object type, object iMsoLegendKey, object autoText);

        /// <summary>
        /// SupportByVersion MSProject 11
        /// </summary>
        /// <param name="type">optional NetOffice.OfficeApi.Enums.XlDataLabelsType Type = 2</param>
        /// <param name="iMsoLegendKey">optional object iMsoLegendKey</param>
        /// <param name="autoText">optional object autoText</param>
        /// <param name="hasLeaderLines">optional object hasLeaderLines</param>
        /// <param name="showSeriesName">optional object showSeriesName</param>
        /// <param name="showCategoryName">optional object showCategoryName</param>
        /// <param name="showValue">optional object showValue</param>
        /// <param name="showPercentage">optional object showPercentage</param>
        /// <param name="showBubbleSize">optional object showBubbleSize</param>
        /// <param name="separator">optional object separator</param>
        [SupportByVersion("MSProject", 11)]
        void ApplyDataLabels(object type, object iMsoLegendKey, object autoText, object hasLeaderLines, object showSeriesName, object showCategoryName, object showValue, object showPercentage, object showBubbleSize, object separator);

        /// <summary>
        /// SupportByVersion MSProject 11
        /// </summary>
        [CustomMethod]
        [SupportByVersion("MSProject", 11)]
        void ApplyDataLabels();

        /// <summary>
        /// SupportByVersion MSProject 11
        /// </summary>
        /// <param name="type">optional NetOffice.OfficeApi.Enums.XlDataLabelsType Type = 2</param>
        [CustomMethod]
        [SupportByVersion("MSProject", 11)]
        void ApplyDataLabels(object type);

        /// <summary>
        /// SupportByVersion MSProject 11
        /// </summary>
        /// <param name="type">optional NetOffice.OfficeApi.Enums.XlDataLabelsType Type = 2</param>
        /// <param name="iMsoLegendKey">optional object iMsoLegendKey</param>
        [CustomMethod]
        [SupportByVersion("MSProject", 11)]
        void ApplyDataLabels(object type, object iMsoLegendKey);

        /// <summary>
        /// SupportByVersion MSProject 11
        /// </summary>
        /// <param name="type">optional NetOffice.OfficeApi.Enums.XlDataLabelsType Type = 2</param>
        /// <param name="iMsoLegendKey">optional object iMsoLegendKey</param>
        /// <param name="autoText">optional object autoText</param>
        [CustomMethod]
        [SupportByVersion("MSProject", 11)]
        void ApplyDataLabels(object type, object iMsoLegendKey, object autoText);

        /// <summary>
        /// SupportByVersion MSProject 11
        /// </summary>
        /// <param name="type">optional NetOffice.OfficeApi.Enums.XlDataLabelsType Type = 2</param>
        /// <param name="iMsoLegendKey">optional object iMsoLegendKey</param>
        /// <param name="autoText">optional object autoText</param>
        /// <param name="hasLeaderLines">optional object hasLeaderLines</param>
        [CustomMethod]
        [SupportByVersion("MSProject", 11)]
        void ApplyDataLabels(object type, object iMsoLegendKey, object autoText, object hasLeaderLines);

        /// <summary>
        /// SupportByVersion MSProject 11
        /// </summary>
        /// <param name="type">optional NetOffice.OfficeApi.Enums.XlDataLabelsType Type = 2</param>
        /// <param name="iMsoLegendKey">optional object iMsoLegendKey</param>
        /// <param name="autoText">optional object autoText</param>
        /// <param name="hasLeaderLines">optional object hasLeaderLines</param>
        /// <param name="showSeriesName">optional object showSeriesName</param>
        [CustomMethod]
        [SupportByVersion("MSProject", 11)]
        void ApplyDataLabels(object type, object iMsoLegendKey, object autoText, object hasLeaderLines, object showSeriesName);

        /// <summary>
        /// SupportByVersion MSProject 11
        /// </summary>
        /// <param name="type">optional NetOffice.OfficeApi.Enums.XlDataLabelsType Type = 2</param>
        /// <param name="iMsoLegendKey">optional object iMsoLegendKey</param>
        /// <param name="autoText">optional object autoText</param>
        /// <param name="hasLeaderLines">optional object hasLeaderLines</param>
        /// <param name="showSeriesName">optional object showSeriesName</param>
        /// <param name="showCategoryName">optional object showCategoryName</param>
        [CustomMethod]
        [SupportByVersion("MSProject", 11)]
        void ApplyDataLabels(object type, object iMsoLegendKey, object autoText, object hasLeaderLines, object showSeriesName, object showCategoryName);

        /// <summary>
        /// SupportByVersion MSProject 11
        /// </summary>
        /// <param name="type">optional NetOffice.OfficeApi.Enums.XlDataLabelsType Type = 2</param>
        /// <param name="iMsoLegendKey">optional object iMsoLegendKey</param>
        /// <param name="autoText">optional object autoText</param>
        /// <param name="hasLeaderLines">optional object hasLeaderLines</param>
        /// <param name="showSeriesName">optional object showSeriesName</param>
        /// <param name="showCategoryName">optional object showCategoryName</param>
        /// <param name="showValue">optional object showValue</param>
        [CustomMethod]
        [SupportByVersion("MSProject", 11)]
        void ApplyDataLabels(object type, object iMsoLegendKey, object autoText, object hasLeaderLines, object showSeriesName, object showCategoryName, object showValue);

        /// <summary>
        /// SupportByVersion MSProject 11
        /// </summary>
        /// <param name="type">optional NetOffice.OfficeApi.Enums.XlDataLabelsType Type = 2</param>
        /// <param name="iMsoLegendKey">optional object iMsoLegendKey</param>
        /// <param name="autoText">optional object autoText</param>
        /// <param name="hasLeaderLines">optional object hasLeaderLines</param>
        /// <param name="showSeriesName">optional object showSeriesName</param>
        /// <param name="showCategoryName">optional object showCategoryName</param>
        /// <param name="showValue">optional object showValue</param>
        /// <param name="showPercentage">optional object showPercentage</param>
        [CustomMethod]
        [SupportByVersion("MSProject", 11)]
        void ApplyDataLabels(object type, object iMsoLegendKey, object autoText, object hasLeaderLines, object showSeriesName, object showCategoryName, object showValue, object showPercentage);

        /// <summary>
        /// SupportByVersion MSProject 11
        /// </summary>
        /// <param name="type">optional NetOffice.OfficeApi.Enums.XlDataLabelsType Type = 2</param>
        /// <param name="iMsoLegendKey">optional object iMsoLegendKey</param>
        /// <param name="autoText">optional object autoText</param>
        /// <param name="hasLeaderLines">optional object hasLeaderLines</param>
        /// <param name="showSeriesName">optional object showSeriesName</param>
        /// <param name="showCategoryName">optional object showCategoryName</param>
        /// <param name="showValue">optional object showValue</param>
        /// <param name="showPercentage">optional object showPercentage</param>
        /// <param name="showBubbleSize">optional object showBubbleSize</param>
        [CustomMethod]
        [SupportByVersion("MSProject", 11)]
        void ApplyDataLabels(object type, object iMsoLegendKey, object autoText, object hasLeaderLines, object showSeriesName, object showCategoryName, object showValue, object showPercentage, object showBubbleSize);

        /// <summary>
        /// SupportByVersion MSProject 11
        /// </summary>
        /// <param name="chartType">NetOffice.OfficeApi.Enums.XlChartType chartType</param>
        /// <param name="typeName">optional object typeName</param>
        [SupportByVersion("MSProject", 11)]
        void ApplyCustomType(NetOffice.OfficeApi.Enums.XlChartType chartType, object typeName);

        /// <summary>
        /// SupportByVersion MSProject 11
        /// </summary>
        /// <param name="chartType">NetOffice.OfficeApi.Enums.XlChartType chartType</param>
        [CustomMethod]
        [SupportByVersion("MSProject", 11)]
        void ApplyCustomType(NetOffice.OfficeApi.Enums.XlChartType chartType);

        /// <summary>
        /// SupportByVersion MSProject 11
        /// </summary>
        /// <param name="x">Int32 x</param>
        /// <param name="y">Int32 y</param>
        /// <param name="elementID">Int32 elementID</param>
        /// <param name="arg1">Int32 arg1</param>
        /// <param name="arg2">Int32 arg2</param>
        [SupportByVersion("MSProject", 11)]
        void GetChartElement(Int32 x, Int32 y, Int32 elementID, Int32 arg1, Int32 arg2);

        /// <summary>
        /// SupportByVersion MSProject 11
        /// </summary>
        /// <param name="source">string source</param>
        /// <param name="plotBy">optional object plotBy</param>
        [SupportByVersion("MSProject", 11)]
        void SetSourceData(string source, object plotBy);

        /// <summary>
        /// SupportByVersion MSProject 11
        /// </summary>
        /// <param name="source">string source</param>
        [CustomMethod]
        [SupportByVersion("MSProject", 11)]
        void SetSourceData(string source);

        /// <summary>
        /// SupportByVersion MSProject 11
        /// </summary>
        /// <param name="type">optional object type</param>
        /// <param name="axisGroup">optional NetOffice.OfficeApi.Enums.XlAxisGroup AxisGroup = 1</param>
        [SupportByVersion("MSProject", 11)]
        object Axes(object type, object axisGroup);

        /// <summary>
        /// SupportByVersion MSProject 11
        /// </summary>
        [CustomMethod]
        [SupportByVersion("MSProject", 11)]
        object Axes();

        /// <summary>
        /// SupportByVersion MSProject 11
        /// </summary>
        /// <param name="type">optional object type</param>
        [CustomMethod]
        [SupportByVersion("MSProject", 11)]
        object Axes(object type);

        /// <summary>
        /// SupportByVersion MSProject 11
        /// </summary>
        /// <param name="rGallery">Int32 rGallery</param>
        /// <param name="varFormat">optional object varFormat</param>
        [SupportByVersion("MSProject", 11)]
        void AutoFormat(Int32 rGallery, object varFormat);

        /// <summary>
        /// SupportByVersion MSProject 11
        /// </summary>
        /// <param name="rGallery">Int32 rGallery</param>
        [CustomMethod]
        [SupportByVersion("MSProject", 11)]
        void AutoFormat(Int32 rGallery);

        /// <summary>
        /// SupportByVersion MSProject 11
        /// </summary>
        /// <param name="bstr">string bstr</param>
        [SupportByVersion("MSProject", 11)]
        void SetBackgroundPicture(string bstr);

        /// <summary>
        /// SupportByVersion MSProject 11
        /// </summary>
        /// <param name="varSource">optional object varSource</param>
        /// <param name="varGallery">optional object varGallery</param>
        /// <param name="varFormat">optional object varFormat</param>
        /// <param name="varPlotBy">optional object varPlotBy</param>
        /// <param name="varCategoryLabels">optional object varCategoryLabels</param>
        /// <param name="varSeriesLabels">optional object varSeriesLabels</param>
        /// <param name="varHasLegend">optional object varHasLegend</param>
        /// <param name="varTitle">optional object varTitle</param>
        /// <param name="varCategoryTitle">optional object varCategoryTitle</param>
        /// <param name="varValueTitle">optional object varValueTitle</param>
        /// <param name="varExtraTitle">optional object varExtraTitle</param>
        [SupportByVersion("MSProject", 11)]
        void ChartWizard(object varSource, object varGallery, object varFormat, object varPlotBy, object varCategoryLabels, object varSeriesLabels, object varHasLegend, object varTitle, object varCategoryTitle, object varValueTitle, object varExtraTitle);

        /// <summary>
        /// SupportByVersion MSProject 11
        /// </summary>
        [CustomMethod]
        [SupportByVersion("MSProject", 11)]
        void ChartWizard();

        /// <summary>
        /// SupportByVersion MSProject 11
        /// </summary>
        /// <param name="varSource">optional object varSource</param>
        [CustomMethod]
        [SupportByVersion("MSProject", 11)]
        void ChartWizard(object varSource);

        /// <summary>
        /// SupportByVersion MSProject 11
        /// </summary>
        /// <param name="varSource">optional object varSource</param>
        /// <param name="varGallery">optional object varGallery</param>
        [CustomMethod]
        [SupportByVersion("MSProject", 11)]
        void ChartWizard(object varSource, object varGallery);

        /// <summary>
        /// SupportByVersion MSProject 11
        /// </summary>
        /// <param name="varSource">optional object varSource</param>
        /// <param name="varGallery">optional object varGallery</param>
        /// <param name="varFormat">optional object varFormat</param>
        [CustomMethod]
        [SupportByVersion("MSProject", 11)]
        void ChartWizard(object varSource, object varGallery, object varFormat);

        /// <summary>
        /// SupportByVersion MSProject 11
        /// </summary>
        /// <param name="varSource">optional object varSource</param>
        /// <param name="varGallery">optional object varGallery</param>
        /// <param name="varFormat">optional object varFormat</param>
        /// <param name="varPlotBy">optional object varPlotBy</param>
        [CustomMethod]
        [SupportByVersion("MSProject", 11)]
        void ChartWizard(object varSource, object varGallery, object varFormat, object varPlotBy);

        /// <summary>
        /// SupportByVersion MSProject 11
        /// </summary>
        /// <param name="varSource">optional object varSource</param>
        /// <param name="varGallery">optional object varGallery</param>
        /// <param name="varFormat">optional object varFormat</param>
        /// <param name="varPlotBy">optional object varPlotBy</param>
        /// <param name="varCategoryLabels">optional object varCategoryLabels</param>
        [CustomMethod]
        [SupportByVersion("MSProject", 11)]
        void ChartWizard(object varSource, object varGallery, object varFormat, object varPlotBy, object varCategoryLabels);

        /// <summary>
        /// SupportByVersion MSProject 11
        /// </summary>
        /// <param name="varSource">optional object varSource</param>
        /// <param name="varGallery">optional object varGallery</param>
        /// <param name="varFormat">optional object varFormat</param>
        /// <param name="varPlotBy">optional object varPlotBy</param>
        /// <param name="varCategoryLabels">optional object varCategoryLabels</param>
        /// <param name="varSeriesLabels">optional object varSeriesLabels</param>
        [CustomMethod]
        [SupportByVersion("MSProject", 11)]
        void ChartWizard(object varSource, object varGallery, object varFormat, object varPlotBy, object varCategoryLabels, object varSeriesLabels);

        /// <summary>
        /// SupportByVersion MSProject 11
        /// </summary>
        /// <param name="varSource">optional object varSource</param>
        /// <param name="varGallery">optional object varGallery</param>
        /// <param name="varFormat">optional object varFormat</param>
        /// <param name="varPlotBy">optional object varPlotBy</param>
        /// <param name="varCategoryLabels">optional object varCategoryLabels</param>
        /// <param name="varSeriesLabels">optional object varSeriesLabels</param>
        /// <param name="varHasLegend">optional object varHasLegend</param>
        [CustomMethod]
        [SupportByVersion("MSProject", 11)]
        void ChartWizard(object varSource, object varGallery, object varFormat, object varPlotBy, object varCategoryLabels, object varSeriesLabels, object varHasLegend);

        /// <summary>
        /// SupportByVersion MSProject 11
        /// </summary>
        /// <param name="varSource">optional object varSource</param>
        /// <param name="varGallery">optional object varGallery</param>
        /// <param name="varFormat">optional object varFormat</param>
        /// <param name="varPlotBy">optional object varPlotBy</param>
        /// <param name="varCategoryLabels">optional object varCategoryLabels</param>
        /// <param name="varSeriesLabels">optional object varSeriesLabels</param>
        /// <param name="varHasLegend">optional object varHasLegend</param>
        /// <param name="varTitle">optional object varTitle</param>
        [CustomMethod]
        [SupportByVersion("MSProject", 11)]
        void ChartWizard(object varSource, object varGallery, object varFormat, object varPlotBy, object varCategoryLabels, object varSeriesLabels, object varHasLegend, object varTitle);

        /// <summary>
        /// SupportByVersion MSProject 11
        /// </summary>
        /// <param name="varSource">optional object varSource</param>
        /// <param name="varGallery">optional object varGallery</param>
        /// <param name="varFormat">optional object varFormat</param>
        /// <param name="varPlotBy">optional object varPlotBy</param>
        /// <param name="varCategoryLabels">optional object varCategoryLabels</param>
        /// <param name="varSeriesLabels">optional object varSeriesLabels</param>
        /// <param name="varHasLegend">optional object varHasLegend</param>
        /// <param name="varTitle">optional object varTitle</param>
        /// <param name="varCategoryTitle">optional object varCategoryTitle</param>
        [CustomMethod]
        [SupportByVersion("MSProject", 11)]
        void ChartWizard(object varSource, object varGallery, object varFormat, object varPlotBy, object varCategoryLabels, object varSeriesLabels, object varHasLegend, object varTitle, object varCategoryTitle);

        /// <summary>
        /// SupportByVersion MSProject 11
        /// </summary>
        /// <param name="varSource">optional object varSource</param>
        /// <param name="varGallery">optional object varGallery</param>
        /// <param name="varFormat">optional object varFormat</param>
        /// <param name="varPlotBy">optional object varPlotBy</param>
        /// <param name="varCategoryLabels">optional object varCategoryLabels</param>
        /// <param name="varSeriesLabels">optional object varSeriesLabels</param>
        /// <param name="varHasLegend">optional object varHasLegend</param>
        /// <param name="varTitle">optional object varTitle</param>
        /// <param name="varCategoryTitle">optional object varCategoryTitle</param>
        /// <param name="varValueTitle">optional object varValueTitle</param>
        [CustomMethod]
        [SupportByVersion("MSProject", 11)]
        void ChartWizard(object varSource, object varGallery, object varFormat, object varPlotBy, object varCategoryLabels, object varSeriesLabels, object varHasLegend, object varTitle, object varCategoryTitle, object varValueTitle);

        /// <summary>
        /// SupportByVersion MSProject 11
        /// </summary>
        /// <param name="appearance">optional Int32 Appearance = 1</param>
        /// <param name="format">optional Int32 Format = -4147</param>
        /// <param name="size">optional Int32 Size = 2</param>
        [SupportByVersion("MSProject", 11)]
        void CopyPicture(object appearance, object format, object size);

        /// <summary>
        /// SupportByVersion MSProject 11
        /// </summary>
        [CustomMethod]
        [SupportByVersion("MSProject", 11)]
        void CopyPicture();

        /// <summary>
        /// SupportByVersion MSProject 11
        /// </summary>
        /// <param name="appearance">optional Int32 Appearance = 1</param>
        [CustomMethod]
        [SupportByVersion("MSProject", 11)]
        void CopyPicture(object appearance);

        /// <summary>
        /// SupportByVersion MSProject 11
        /// </summary>
        /// <param name="appearance">optional Int32 Appearance = 1</param>
        /// <param name="format">optional Int32 Format = -4147</param>
        [CustomMethod]
        [SupportByVersion("MSProject", 11)]
        void CopyPicture(object appearance, object format);

        /// <summary>
        /// SupportByVersion MSProject 11
        /// </summary>
        /// <param name="varName">object varName</param>
        /// <param name="localeID">Int32 localeID</param>
        /// <param name="objType">Int32 objType</param>
        [SupportByVersion("MSProject", 11)]
        object Evaluate(object varName, Int32 localeID, out Int32 objType);

		/// <summary>
		/// SupportByVersion MSProject 11
		/// </summary>
		/// <param name="varName">object varName</param>
		/// <param name="localeID">Int32 localeID</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [SupportByVersion("MSProject", 11)]
        object _Evaluate(object varName, Int32 localeID);

        /// <summary>
        /// SupportByVersion MSProject 11
        /// </summary>
        /// <param name="varType">optional object varType</param>
        [SupportByVersion("MSProject", 11)]
        void Paste(object varType);

        /// <summary>
        /// SupportByVersion MSProject 11
        /// </summary>
        [CustomMethod]
        [SupportByVersion("MSProject", 11)]
        void Paste();

        /// <summary>
        /// SupportByVersion MSProject 11
        /// </summary>
        /// <param name="bstr">string bstr</param>
        /// <param name="varFilterName">optional object varFilterName</param>
        /// <param name="varInteractive">optional object varInteractive</param>
        [SupportByVersion("MSProject", 11)]
        bool Export(string bstr, object varFilterName, object varInteractive);

        /// <summary>
        /// SupportByVersion MSProject 11
        /// </summary>
        /// <param name="bstr">string bstr</param>
        [CustomMethod]
        [SupportByVersion("MSProject", 11)]
        bool Export(string bstr);

        /// <summary>
        /// SupportByVersion MSProject 11
        /// </summary>
        /// <param name="bstr">string bstr</param>
        /// <param name="varFilterName">optional object varFilterName</param>
        [CustomMethod]
        [SupportByVersion("MSProject", 11)]
        bool Export(string bstr, object varFilterName);

        /// <summary>
        /// SupportByVersion MSProject 11
        /// </summary>
        /// <param name="varName">object varName</param>
        [SupportByVersion("MSProject", 11)]
        void SetDefaultChart(object varName);

        /// <summary>
        /// SupportByVersion MSProject 11
        /// </summary>
        /// <param name="bstrFileName">string bstrFileName</param>
        [SupportByVersion("MSProject", 11)]
        void ApplyChartTemplate(string bstrFileName);

        /// <summary>
        /// SupportByVersion MSProject 11
        /// </summary>
        /// <param name="bstrFileName">string bstrFileName</param>
        [SupportByVersion("MSProject", 11)]
        void SaveChartTemplate(string bstrFileName);

        /// <summary>
        /// SupportByVersion MSProject 11
        /// </summary>
        [SupportByVersion("MSProject", 11)]
        void ClearToMatchStyle();

        /// <summary>
        /// SupportByVersion MSProject 11
        /// </summary>
        [SupportByVersion("MSProject", 11)]
        void RefreshPivotTable();

        /// <summary>
        /// SupportByVersion MSProject 11
        /// </summary>
        /// <param name="layout">Int32 layout</param>
        /// <param name="varChartType">optional object varChartType</param>
        [SupportByVersion("MSProject", 11)]
        void ApplyLayout(Int32 layout, object varChartType);

        /// <summary>
        /// SupportByVersion MSProject 11
        /// </summary>
        /// <param name="layout">Int32 layout</param>
        [CustomMethod]
        [SupportByVersion("MSProject", 11)]
        void ApplyLayout(Int32 layout);

        /// <summary>
        /// SupportByVersion MSProject 11
        /// </summary>
        [SupportByVersion("MSProject", 11)]
        void Refresh();

        /// <summary>
        /// SupportByVersion MSProject 11
        /// </summary>
        /// <param name="rHS">NetOffice.OfficeApi.Enums.MsoChartElementType rHS</param>
        [SupportByVersion("MSProject", 11)]
        void SetElement(NetOffice.OfficeApi.Enums.MsoChartElementType rHS);

        /// <summary>
        /// SupportByVersion MSProject 11
        /// </summary>
        /// <param name="index">optional object index</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [SupportByVersion("MSProject", 11)]
        object AreaGroups(object index);

        /// <summary>
        /// SupportByVersion MSProject 11
        /// </summary>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [CustomMethod]
        [SupportByVersion("MSProject", 11)]
        object AreaGroups();

        /// <summary>
        /// SupportByVersion MSProject 11
        /// </summary>
        /// <param name="index">optional object index</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [SupportByVersion("MSProject", 11)]
        object BarGroups(object index);

        /// <summary>
        /// SupportByVersion MSProject 11
        /// </summary>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [CustomMethod]
        [SupportByVersion("MSProject", 11)]
        object BarGroups();

        /// <summary>
        /// SupportByVersion MSProject 11
        /// </summary>
        /// <param name="index">optional object index</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [SupportByVersion("MSProject", 11)]
        object ColumnGroups(object index);

        /// <summary>
        /// SupportByVersion MSProject 11
        /// </summary>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [CustomMethod]
        [SupportByVersion("MSProject", 11)]
        object ColumnGroups();

        /// <summary>
        /// SupportByVersion MSProject 11
        /// </summary>
        /// <param name="index">optional object index</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [SupportByVersion("MSProject", 11)]
        object LineGroups(object index);

        /// <summary>
        /// SupportByVersion MSProject 11
        /// </summary>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [CustomMethod]
        [SupportByVersion("MSProject", 11)]
        object LineGroups();

        /// <summary>
        /// SupportByVersion MSProject 11
        /// </summary>
        /// <param name="index">optional object index</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [SupportByVersion("MSProject", 11)]
        object PieGroups(object index);

        /// <summary>
        /// SupportByVersion MSProject 11
        /// </summary>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [CustomMethod]
        [SupportByVersion("MSProject", 11)]
        object PieGroups();

        /// <summary>
        /// SupportByVersion MSProject 11
        /// </summary>
        /// <param name="index">optional object index</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [SupportByVersion("MSProject", 11)]
        object DoughnutGroups(object index);

        /// <summary>
        /// SupportByVersion MSProject 11
        /// </summary>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [CustomMethod]
        [SupportByVersion("MSProject", 11)]
        object DoughnutGroups();

        /// <summary>
        /// SupportByVersion MSProject 11
        /// </summary>
        /// <param name="index">optional object index</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [SupportByVersion("MSProject", 11)]
        object RadarGroups(object index);

        /// <summary>
        /// SupportByVersion MSProject 11
        /// </summary>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [CustomMethod]
        [SupportByVersion("MSProject", 11)]
        object RadarGroups();

        /// <summary>
        /// SupportByVersion MSProject 11
        /// </summary>
        /// <param name="index">optional object index</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [SupportByVersion("MSProject", 11)]
        object XYGroups(object index);

        /// <summary>
        /// SupportByVersion MSProject 11
        /// </summary>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [CustomMethod]
        [SupportByVersion("MSProject", 11)]
        object XYGroups();

        /// <summary>
        /// SupportByVersion MSProject 11
        /// </summary>
        [SupportByVersion("MSProject", 11)]
        object Delete();

        /// <summary>
        /// SupportByVersion MSProject 11
        /// </summary>
        [SupportByVersion("MSProject", 11)]
        object Copy();

        /// <summary>
        /// SupportByVersion MSProject 11
        /// </summary>
        /// <param name="replace">optional object replace</param>
        [SupportByVersion("MSProject", 11)]
        object Select(object replace);

        /// <summary>
        /// SupportByVersion MSProject 11
        /// </summary>
        [CustomMethod]
        [SupportByVersion("MSProject", 11)]
        object Select();

        /// <summary>
        /// SupportByVersion MSProject 11
        /// </summary>
        /// <param name="task">bool task</param>
        /// <param name="timephased">bool timephased</param>
        /// <param name="groupName">optional string GroupName = </param>
        /// <param name="filterName">optional string FilterName = </param>
        /// <param name="labelField">optional NetOffice.MSProjectApi.Enums.PjField LabelField = -1</param>
        /// <param name="outlineLevel">optional Int32 OutlineLevel = -1</param>
        /// <param name="safeArrayOfPjField">optional object safeArrayOfPjField</param>
        /// <param name="safeArrayOfPjTimescaledData">optional object safeArrayOfPjTimescaledData</param>
        /// <param name="timeScaleUnit">optional NetOffice.MSProjectApi.Enums.PjTimescaleUnit TimeScaleUnit = 4</param>
        /// <param name="timescaleUnitCount">optional Int32 TimescaleUnitCount = 1</param>
        /// <param name="startDate">optional object startDate</param>
        /// <param name="finishDate">optional object finishDate</param>
        [SupportByVersion("MSProject", 11)]
        void UpdateChartData(bool task, bool timephased, object groupName, object filterName, object labelField, object outlineLevel, object safeArrayOfPjField, object safeArrayOfPjTimescaledData, object timeScaleUnit, object timescaleUnitCount, object startDate, object finishDate);

        /// <summary>
        /// SupportByVersion MSProject 11
        /// </summary>
        /// <param name="task">bool task</param>
        /// <param name="timephased">bool timephased</param>
        [CustomMethod]
        [SupportByVersion("MSProject", 11)]
        void UpdateChartData(bool task, bool timephased);

        /// <summary>
        /// SupportByVersion MSProject 11
        /// </summary>
        /// <param name="task">bool task</param>
        /// <param name="timephased">bool timephased</param>
        /// <param name="groupName">optional string GroupName = </param>
        [CustomMethod]
        [SupportByVersion("MSProject", 11)]
        void UpdateChartData(bool task, bool timephased, object groupName);

        /// <summary>
        /// SupportByVersion MSProject 11
        /// </summary>
        /// <param name="task">bool task</param>
        /// <param name="timephased">bool timephased</param>
        /// <param name="groupName">optional string GroupName = </param>
        /// <param name="filterName">optional string FilterName = </param>
        [CustomMethod]
        [SupportByVersion("MSProject", 11)]
        void UpdateChartData(bool task, bool timephased, object groupName, object filterName);

        /// <summary>
        /// SupportByVersion MSProject 11
        /// </summary>
        /// <param name="task">bool task</param>
        /// <param name="timephased">bool timephased</param>
        /// <param name="groupName">optional string GroupName = </param>
        /// <param name="filterName">optional string FilterName = </param>
        /// <param name="labelField">optional NetOffice.MSProjectApi.Enums.PjField LabelField = -1</param>
        [CustomMethod]
        [SupportByVersion("MSProject", 11)]
        void UpdateChartData(bool task, bool timephased, object groupName, object filterName, object labelField);

        /// <summary>
        /// SupportByVersion MSProject 11
        /// </summary>
        /// <param name="task">bool task</param>
        /// <param name="timephased">bool timephased</param>
        /// <param name="groupName">optional string GroupName = </param>
        /// <param name="filterName">optional string FilterName = </param>
        /// <param name="labelField">optional NetOffice.MSProjectApi.Enums.PjField LabelField = -1</param>
        /// <param name="outlineLevel">optional Int32 OutlineLevel = -1</param>
        [CustomMethod]
        [SupportByVersion("MSProject", 11)]
        void UpdateChartData(bool task, bool timephased, object groupName, object filterName, object labelField, object outlineLevel);

        /// <summary>
        /// SupportByVersion MSProject 11
        /// </summary>
        /// <param name="task">bool task</param>
        /// <param name="timephased">bool timephased</param>
        /// <param name="groupName">optional string GroupName = </param>
        /// <param name="filterName">optional string FilterName = </param>
        /// <param name="labelField">optional NetOffice.MSProjectApi.Enums.PjField LabelField = -1</param>
        /// <param name="outlineLevel">optional Int32 OutlineLevel = -1</param>
        /// <param name="safeArrayOfPjField">optional object safeArrayOfPjField</param>
        [CustomMethod]
        [SupportByVersion("MSProject", 11)]
        void UpdateChartData(bool task, bool timephased, object groupName, object filterName, object labelField, object outlineLevel, object safeArrayOfPjField);

        /// <summary>
        /// SupportByVersion MSProject 11
        /// </summary>
        /// <param name="task">bool task</param>
        /// <param name="timephased">bool timephased</param>
        /// <param name="groupName">optional string GroupName = </param>
        /// <param name="filterName">optional string FilterName = </param>
        /// <param name="labelField">optional NetOffice.MSProjectApi.Enums.PjField LabelField = -1</param>
        /// <param name="outlineLevel">optional Int32 OutlineLevel = -1</param>
        /// <param name="safeArrayOfPjField">optional object safeArrayOfPjField</param>
        /// <param name="safeArrayOfPjTimescaledData">optional object safeArrayOfPjTimescaledData</param>
        [CustomMethod]
        [SupportByVersion("MSProject", 11)]
        void UpdateChartData(bool task, bool timephased, object groupName, object filterName, object labelField, object outlineLevel, object safeArrayOfPjField, object safeArrayOfPjTimescaledData);

        /// <summary>
        /// SupportByVersion MSProject 11
        /// </summary>
        /// <param name="task">bool task</param>
        /// <param name="timephased">bool timephased</param>
        /// <param name="groupName">optional string GroupName = </param>
        /// <param name="filterName">optional string FilterName = </param>
        /// <param name="labelField">optional NetOffice.MSProjectApi.Enums.PjField LabelField = -1</param>
        /// <param name="outlineLevel">optional Int32 OutlineLevel = -1</param>
        /// <param name="safeArrayOfPjField">optional object safeArrayOfPjField</param>
        /// <param name="safeArrayOfPjTimescaledData">optional object safeArrayOfPjTimescaledData</param>
        /// <param name="timeScaleUnit">optional NetOffice.MSProjectApi.Enums.PjTimescaleUnit TimeScaleUnit = 4</param>
        [CustomMethod]
        [SupportByVersion("MSProject", 11)]
        void UpdateChartData(bool task, bool timephased, object groupName, object filterName, object labelField, object outlineLevel, object safeArrayOfPjField, object safeArrayOfPjTimescaledData, object timeScaleUnit);

        /// <summary>
        /// SupportByVersion MSProject 11
        /// </summary>
        /// <param name="task">bool task</param>
        /// <param name="timephased">bool timephased</param>
        /// <param name="groupName">optional string GroupName = </param>
        /// <param name="filterName">optional string FilterName = </param>
        /// <param name="labelField">optional NetOffice.MSProjectApi.Enums.PjField LabelField = -1</param>
        /// <param name="outlineLevel">optional Int32 OutlineLevel = -1</param>
        /// <param name="safeArrayOfPjField">optional object safeArrayOfPjField</param>
        /// <param name="safeArrayOfPjTimescaledData">optional object safeArrayOfPjTimescaledData</param>
        /// <param name="timeScaleUnit">optional NetOffice.MSProjectApi.Enums.PjTimescaleUnit TimeScaleUnit = 4</param>
        /// <param name="timescaleUnitCount">optional Int32 TimescaleUnitCount = 1</param>
        [CustomMethod]
        [SupportByVersion("MSProject", 11)]
        void UpdateChartData(bool task, bool timephased, object groupName, object filterName, object labelField, object outlineLevel, object safeArrayOfPjField, object safeArrayOfPjTimescaledData, object timeScaleUnit, object timescaleUnitCount);

        /// <summary>
        /// SupportByVersion MSProject 11
        /// </summary>
        /// <param name="task">bool task</param>
        /// <param name="timephased">bool timephased</param>
        /// <param name="groupName">optional string GroupName = </param>
        /// <param name="filterName">optional string FilterName = </param>
        /// <param name="labelField">optional NetOffice.MSProjectApi.Enums.PjField LabelField = -1</param>
        /// <param name="outlineLevel">optional Int32 OutlineLevel = -1</param>
        /// <param name="safeArrayOfPjField">optional object safeArrayOfPjField</param>
        /// <param name="safeArrayOfPjTimescaledData">optional object safeArrayOfPjTimescaledData</param>
        /// <param name="timeScaleUnit">optional NetOffice.MSProjectApi.Enums.PjTimescaleUnit TimeScaleUnit = 4</param>
        /// <param name="timescaleUnitCount">optional Int32 TimescaleUnitCount = 1</param>
        /// <param name="startDate">optional object startDate</param>
        [CustomMethod]
        [SupportByVersion("MSProject", 11)]
        void UpdateChartData(bool task, bool timephased, object groupName, object filterName, object labelField, object outlineLevel, object safeArrayOfPjField, object safeArrayOfPjTimescaledData, object timeScaleUnit, object timescaleUnitCount, object startDate);

        /// <summary>
        /// SupportByVersion MSProject 11
        /// </summary>
        [SupportByVersion("MSProject", 11)]
        void ClearToMatchColorStyle();

        #endregion
    }
}
