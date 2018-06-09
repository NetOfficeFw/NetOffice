using System.Reflection;
using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.OfficeApi
{
    /// <summary>
    /// IMsoChart
    /// </summary>
    [SyntaxBypass]
    public interface IMsoChart_ : ICOMObject
    {
        #region Properties

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// Get
        /// Unknown COM Proxy
        /// </summary>
        /// <param name="pvarIndex">optional object pvarIndex</param>
        /// <param name="varIgallery">optional object varIgallery</param>
        [SupportByVersion("Office", 12, 14, 15, 16), ProxyResult]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        object get_ChartGroups(object pvarIndex, object varIgallery);

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// Alias for get_ChartGroups
        /// Unknown COM Proxy
        /// </summary>
        /// <param name="pvarIndex">optional object pvarIndex</param>
        /// <param name="varIgallery">optional object varIgallery</param>
        [SupportByVersion("Office", 12, 14, 15, 16), ProxyResult, Redirect("get_ChartGroups")]
        object ChartGroups(object pvarIndex, object varIgallery);

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// Get
        /// Unknown COM Proxy
        /// </summary>
        /// <param name="pvarIndex">optional object pvarIndex</param>
        [SupportByVersion("Office", 12, 14, 15, 16), ProxyResult]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        object get_ChartGroups(object pvarIndex);

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// Alias for get_ChartGroups
        /// Unknown COM Proxy
        /// </summary>
        /// <param name="pvarIndex">optional object pvarIndex</param>
        [SupportByVersion("Office", 12, 14, 15, 16), ProxyResult, Redirect("get_ChartGroups")]
        object ChartGroups(object pvarIndex);

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <param name="axisType">optional object axisType</param>
        /// <param name="axisGroup">optional object axisGroup</param>
        [SupportByVersion("Office", 12, 14, 15, 16)]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        object get_HasAxis(object axisType, object axisGroup);

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <param name="axisType">optional object axisType</param>
        /// <param name="axisGroup">optional object axisGroup</param>
        /// <param name="value">optional object value</param>
        [SupportByVersion("Office", 12, 14, 15, 16)]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        void set_HasAxis(object axisType, object axisGroup, object value);

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// Alias for get_HasAxis
        /// </summary>
        /// <param name="axisType">optional object axisType</param>
        /// <param name="axisGroup">optional object axisGroup</param>
        [SupportByVersion("Office", 12, 14, 15, 16), Redirect("get_HasAxis")]
        object HasAxis(object axisType, object axisGroup);

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <param name="axisType">optional object axisType</param>
        [SupportByVersion("Office", 12, 14, 15, 16)]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        object get_HasAxis(object axisType);

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <param name="axisType">optional object axisType</param>
        /// <param name="value">optional object value</param>
        [SupportByVersion("Office", 12, 14, 15, 16)]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        void set_HasAxis(object axisType, object value);

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// Alias for get_HasAxis
        /// </summary>
        /// <param name="axisType">optional object axisType</param>
        [SupportByVersion("Office", 12, 14, 15, 16), Redirect("get_HasAxis")]
        object HasAxis(object axisType);

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <param name="fBackWall">optional bool fBackWall</param>
        [SupportByVersion("Office", 12, 14, 15, 16)]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        NetOffice.OfficeApi.IMsoWalls get_Walls(object fBackWall);

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// Alias for get_Walls
        /// </summary>
        /// <param name="fBackWall">optional bool fBackWall</param>
        [SupportByVersion("Office", 12, 14, 15, 16), Redirect("get_Walls")]
        NetOffice.OfficeApi.IMsoWalls Walls(object fBackWall);

        #endregion
    }

    /// <summary>
    /// DispatchInterface IMsoChart 
    /// SupportByVersion Office, 12,14,15,16
    /// </summary>
    [SupportByVersion("Office", 12, 14, 15, 16)]
    [EntityType(EntityType.IsDispatchInterface)]
	[TypeId("000C1709-0000-0000-C000-000000000046")]
    public interface IMsoChart : IMsoChart_
    {
        #region Properties

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// Get
        /// Unknown COM Proxy
        /// </summary>
        [SupportByVersion("Office", 12, 14, 15, 16), ProxyResult]
        object Parent { get; }

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Office", 12, 14, 15, 16)]
        bool HasTitle { get; set; }

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Office", 12, 14, 15, 16)]
        [BaseResult]
        NetOffice.OfficeApi.IMsoChartTitle ChartTitle { get; }

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Office", 12, 14, 15, 16)]
        Int32 DepthPercent { get; set; }

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Office", 12, 14, 15, 16)]
        Int32 Elevation { get; set; }

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Office", 12, 14, 15, 16)]
        Int32 GapDepth { get; set; }

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Office", 12, 14, 15, 16)]
        Int32 HeightPercent { get; set; }

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Office", 12, 14, 15, 16)]
        Int32 Perspective { get; set; }

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Office", 12, 14, 15, 16)]
        object RightAngleAxes { get; set; }

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Office", 12, 14, 15, 16)]
        object Rotation { get; set; }

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Office", 12, 14, 15, 16)]
        NetOffice.OfficeApi.Enums.XlDisplayBlanksAs DisplayBlanksAs { get; set; }

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Office", 12, 14, 15, 16)]
        bool ProtectData { get; set; }

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Office", 12, 14, 15, 16)]
        bool ProtectFormatting { get; set; }

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Office", 12, 14, 15, 16)]
        bool ProtectGoalSeek { get; set; }

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Office", 12, 14, 15, 16)]
        bool ProtectSelection { get; set; }

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Office", 12, 14, 15, 16)]
        bool ProtectChartObjects { get; set; }

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// Get
        /// Unknown COM Proxy
        /// </summary>
        [SupportByVersion("Office", 12, 14, 15, 16), ProxyResult]
        new object ChartGroups { get; }

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Office", 12, 14, 15, 16)]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        Int32 SubType { get; set; }

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Office", 12, 14, 15, 16)]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        Int32 Type { get; set; }

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Office", 12, 14, 15, 16)]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        NetOffice.OfficeApi.IMsoCorners Corners { get; }

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Office", 12, 14, 15, 16)]
        NetOffice.OfficeApi.Enums.XlChartType ChartType { get; set; }

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Office", 12, 14, 15, 16)]
        bool HasDataTable { get; set; }

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Office", 12, 14, 15, 16)]
        NetOffice.OfficeApi.Enums.XlRowCol PlotBy { get; set; }

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Office", 12, 14, 15, 16)]
        bool HasLegend { get; set; }

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Office", 12, 14, 15, 16)]
        NetOffice.OfficeApi.IMsoLegend Legend { get; }

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Office", 12, 14, 15, 16)]
        new object HasAxis { get; set; }

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Office", 12, 14, 15, 16)]
        new NetOffice.OfficeApi.IMsoWalls Walls { get; }

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Office", 12, 14, 15, 16)]
        NetOffice.OfficeApi.IMsoFloor Floor { get;}

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Office", 12, 14, 15, 16)]
        NetOffice.OfficeApi.IMsoPlotArea PlotArea { get; }

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Office", 12, 14, 15, 16)]
        bool PlotVisibleOnly { get; set; }

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Office", 12, 14, 15, 16)]
        NetOffice.OfficeApi.IMsoChartArea ChartArea { get; }

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Office", 12, 14, 15, 16)]
        bool AutoScaling { get; set; }

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Office", 12, 14, 15, 16)]
        NetOffice.OfficeApi.IMsoDataTable DataTable { get; }

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Office", 12, 14, 15, 16)]
        NetOffice.OfficeApi.Enums.XlBarShape BarShape { get; set; }

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Office", 12, 14, 15, 16)]
        NetOffice.OfficeApi.IMsoWalls SideWall { get; }

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Office", 12, 14, 15, 16)]
        NetOffice.OfficeApi.IMsoWalls BackWall { get; }

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Office", 12, 14, 15, 16)]
        object ChartStyle { get; set; }

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// Get
        /// Unknown COM Proxy
        /// </summary>
        [SupportByVersion("Office", 12, 14, 15, 16), ProxyResult]
        object PivotLayout { get; }

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Office", 12, 14, 15, 16)]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        bool HasPivotFields { get; set; }

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Office", 12, 14, 15, 16)]
        bool ShowDataLabelsOverMaximum { get; set; }

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// Get
        /// Unknown COM Proxy
        /// </summary>
        [SupportByVersion("Office", 12, 14, 15, 16), ProxyResult]
        object Selection { get; }

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Office", 12, 14, 15, 16)]
        NetOffice.OfficeApi.IMsoChartData ChartData { get; }

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Office", 12, 14, 15, 16)]
        NetOffice.OfficeApi.IMsoChartFormat Format { get; }

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Office", 12, 14, 15, 16)]
        NetOffice.OfficeApi.Shapes Shapes { get;}

        /// <summary>
        /// SupportByVersion Office 14, 15, 16
        /// Get
        /// Unknown COM Proxy
        /// </summary>
        [SupportByVersion("Office", 14, 15, 16), ProxyResult]
        object Application { get; }

        /// <summary>
        /// SupportByVersion Office 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Office", 14, 15, 16)]
        Int32 Creator { get; }

        /// <summary>
        /// SupportByVersion Office 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Office", 14, 15, 16)]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        NetOffice.OfficeApi.IMsoChartGroup Area3DGroup { get; }

        /// <summary>
        /// SupportByVersion Office 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Office", 14, 15, 16)]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        NetOffice.OfficeApi.IMsoChartGroup Bar3DGroup { get; }

        /// <summary>
        /// SupportByVersion Office 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Office", 14, 15, 16)]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        NetOffice.OfficeApi.IMsoChartGroup Column3DGroup { get; }

        /// <summary>
        /// SupportByVersion Office 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Office", 14, 15, 16)]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        NetOffice.OfficeApi.IMsoChartGroup Line3DGroup { get; }

        /// <summary>
        /// SupportByVersion Office 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Office", 14, 15, 16)]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        NetOffice.OfficeApi.IMsoChartGroup Pie3DGroup { get; }

        /// <summary>
        /// SupportByVersion Office 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Office", 14, 15, 16)]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        NetOffice.OfficeApi.IMsoChartGroup SurfaceGroup { get; }

        /// <summary>
        /// SupportByVersion Office 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Office", 14, 15, 16)]
        bool ShowReportFilterFieldButtons { get; set; }

        /// <summary>
        /// SupportByVersion Office 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Office", 14, 15, 16)]
        bool ShowLegendFieldButtons { get; set; }

        /// <summary>
        /// SupportByVersion Office 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Office", 14, 15, 16)]
        bool ShowAxisFieldButtons { get; set; }

        /// <summary>
        /// SupportByVersion Office 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Office", 14, 15, 16)]
        bool ShowValueFieldButtons { get; set; }

        /// <summary>
        /// SupportByVersion Office 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Office", 14, 15, 16)]
        bool ShowAllFieldButtons { get; set; }

        /// <summary>
        /// SupportByVersion Office 15,16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Office", 15, 16)]
        bool ProtectChartSheetFormatting { get; set; }

        /// <summary>
        /// SupportByVersion Office 15,16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Office", 15, 16)]
        NetOffice.OfficeApi.Enums.XlCategoryLabelLevel CategoryLabelLevel { get; set; }

        /// <summary>
        /// SupportByVersion Office 15,16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Office", 15, 16)]
        NetOffice.OfficeApi.Enums.XlSeriesNameLevel SeriesNameLevel { get; set; }

        /// <summary>
        /// SupportByVersion Office 15,16
        /// Get
        /// </summary>
        [SupportByVersion("Office", 15, 16)]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        bool HasHiddenContent { get; }

        /// <summary>
        /// SupportByVersion Office 15,16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Office", 15, 16)]
        object ChartColor { get; set; }

        #endregion

        #region Methods

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// </summary>
        /// <param name="password">optional object password</param>
        [SupportByVersion("Office", 12, 14, 15, 16)]
        void UnProtect(object password);

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// </summary>
        [CustomMethod]
        [SupportByVersion("Office", 12, 14, 15, 16)]
        void UnProtect();

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// </summary>
        /// <param name="password">optional object password</param>
        /// <param name="drawingObjects">optional object drawingObjects</param>
        /// <param name="contents">optional object contents</param>
        /// <param name="scenarios">optional object scenarios</param>
        /// <param name="userInterfaceOnly">optional object userInterfaceOnly</param>
        [SupportByVersion("Office", 12, 14, 15, 16)]
        void Protect(object password, object drawingObjects, object contents, object scenarios, object userInterfaceOnly);

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// </summary>
        [CustomMethod]
        [SupportByVersion("Office", 12, 14, 15, 16)]
        void Protect();

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// </summary>
        /// <param name="password">optional object password</param>
        [CustomMethod]
        [SupportByVersion("Office", 12, 14, 15, 16)]
        void Protect(object password);

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// </summary>
        /// <param name="password">optional object password</param>
        /// <param name="drawingObjects">optional object drawingObjects</param>
        [CustomMethod]
        [SupportByVersion("Office", 12, 14, 15, 16)]
        void Protect(object password, object drawingObjects);

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// </summary>
        /// <param name="password">optional object password</param>
        /// <param name="drawingObjects">optional object drawingObjects</param>
        /// <param name="contents">optional object contents</param>
        [CustomMethod]
        [SupportByVersion("Office", 12, 14, 15, 16)]
        void Protect(object password, object drawingObjects, object contents);

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// </summary>
        /// <param name="password">optional object password</param>
        /// <param name="drawingObjects">optional object drawingObjects</param>
        /// <param name="contents">optional object contents</param>
        /// <param name="scenarios">optional object scenarios</param>
        [CustomMethod]
        [SupportByVersion("Office", 12, 14, 15, 16)]
        void Protect(object password, object drawingObjects, object contents, object scenarios);

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// </summary>
        /// <param name="index">optional object index</param>
        [SupportByVersion("Office", 12, 14, 15, 16)]
        object SeriesCollection(object index);

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// </summary>
        [CustomMethod]
        [SupportByVersion("Office", 12, 14, 15, 16)]
        object SeriesCollection();

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// </summary>
        /// <param name="type">optional NetOffice.OfficeApi.Enums.XlDataLabelsType Type = 2</param>
        /// <param name="iMsoLegendKey">optional object iMsoLegendKey</param>
        /// <param name="autoText">optional object autoText</param>
        /// <param name="hasLeaderLines">optional object hasLeaderLines</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [SupportByVersion("Office", 12, 14, 15, 16)]
        void _ApplyDataLabels(object type, object iMsoLegendKey, object autoText, object hasLeaderLines);

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// </summary>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [CustomMethod]
        [SupportByVersion("Office", 12, 14, 15, 16)]
        void _ApplyDataLabels();

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// </summary>
        /// <param name="type">optional NetOffice.OfficeApi.Enums.XlDataLabelsType Type = 2</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [CustomMethod]
        [SupportByVersion("Office", 12, 14, 15, 16)]
        void _ApplyDataLabels(object type);

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// </summary>
        /// <param name="type">optional NetOffice.OfficeApi.Enums.XlDataLabelsType Type = 2</param>
        /// <param name="iMsoLegendKey">optional object iMsoLegendKey</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [CustomMethod]
        [SupportByVersion("Office", 12, 14, 15, 16)]
        void _ApplyDataLabels(object type, object iMsoLegendKey);

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// </summary>
        /// <param name="type">optional NetOffice.OfficeApi.Enums.XlDataLabelsType Type = 2</param>
        /// <param name="iMsoLegendKey">optional object iMsoLegendKey</param>
        /// <param name="autoText">optional object autoText</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [CustomMethod]
        [SupportByVersion("Office", 12, 14, 15, 16)]
        void _ApplyDataLabels(object type, object iMsoLegendKey, object autoText);

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
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
        [SupportByVersion("Office", 12, 14, 15, 16)]
        void ApplyDataLabels(object type, object iMsoLegendKey, object autoText, object hasLeaderLines, object showSeriesName, object showCategoryName, object showValue, object showPercentage, object showBubbleSize, object separator);

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// </summary>
        [CustomMethod]
        [SupportByVersion("Office", 12, 14, 15, 16)]
        void ApplyDataLabels();

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// </summary>
        /// <param name="type">optional NetOffice.OfficeApi.Enums.XlDataLabelsType Type = 2</param>
        [CustomMethod]
        [SupportByVersion("Office", 12, 14, 15, 16)]
        void ApplyDataLabels(object type);

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// </summary>
        /// <param name="type">optional NetOffice.OfficeApi.Enums.XlDataLabelsType Type = 2</param>
        /// <param name="iMsoLegendKey">optional object iMsoLegendKey</param>
        [CustomMethod]
        [SupportByVersion("Office", 12, 14, 15, 16)]
        void ApplyDataLabels(object type, object iMsoLegendKey);

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// </summary>
        /// <param name="type">optional NetOffice.OfficeApi.Enums.XlDataLabelsType Type = 2</param>
        /// <param name="iMsoLegendKey">optional object iMsoLegendKey</param>
        /// <param name="autoText">optional object autoText</param>
        [CustomMethod]
        [SupportByVersion("Office", 12, 14, 15, 16)]
        void ApplyDataLabels(object type, object iMsoLegendKey, object autoText);

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// </summary>
        /// <param name="type">optional NetOffice.OfficeApi.Enums.XlDataLabelsType Type = 2</param>
        /// <param name="iMsoLegendKey">optional object iMsoLegendKey</param>
        /// <param name="autoText">optional object autoText</param>
        /// <param name="hasLeaderLines">optional object hasLeaderLines</param>
        [CustomMethod]
        [SupportByVersion("Office", 12, 14, 15, 16)]
        void ApplyDataLabels(object type, object iMsoLegendKey, object autoText, object hasLeaderLines);

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// </summary>
        /// <param name="type">optional NetOffice.OfficeApi.Enums.XlDataLabelsType Type = 2</param>
        /// <param name="iMsoLegendKey">optional object iMsoLegendKey</param>
        /// <param name="autoText">optional object autoText</param>
        /// <param name="hasLeaderLines">optional object hasLeaderLines</param>
        /// <param name="showSeriesName">optional object showSeriesName</param>
        [CustomMethod]
        [SupportByVersion("Office", 12, 14, 15, 16)]
        void ApplyDataLabels(object type, object iMsoLegendKey, object autoText, object hasLeaderLines, object showSeriesName);

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// </summary>
        /// <param name="type">optional NetOffice.OfficeApi.Enums.XlDataLabelsType Type = 2</param>
        /// <param name="iMsoLegendKey">optional object iMsoLegendKey</param>
        /// <param name="autoText">optional object autoText</param>
        /// <param name="hasLeaderLines">optional object hasLeaderLines</param>
        /// <param name="showSeriesName">optional object showSeriesName</param>
        /// <param name="showCategoryName">optional object showCategoryName</param>
        [CustomMethod]
        [SupportByVersion("Office", 12, 14, 15, 16)]
        void ApplyDataLabels(object type, object iMsoLegendKey, object autoText, object hasLeaderLines, object showSeriesName, object showCategoryName);

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// </summary>
        /// <param name="type">optional NetOffice.OfficeApi.Enums.XlDataLabelsType Type = 2</param>
        /// <param name="iMsoLegendKey">optional object iMsoLegendKey</param>
        /// <param name="autoText">optional object autoText</param>
        /// <param name="hasLeaderLines">optional object hasLeaderLines</param>
        /// <param name="showSeriesName">optional object showSeriesName</param>
        /// <param name="showCategoryName">optional object showCategoryName</param>
        /// <param name="showValue">optional object showValue</param>
        [CustomMethod]
        [SupportByVersion("Office", 12, 14, 15, 16)]
        void ApplyDataLabels(object type, object iMsoLegendKey, object autoText, object hasLeaderLines, object showSeriesName, object showCategoryName, object showValue);

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
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
        [SupportByVersion("Office", 12, 14, 15, 16)]
        void ApplyDataLabels(object type, object iMsoLegendKey, object autoText, object hasLeaderLines, object showSeriesName, object showCategoryName, object showValue, object showPercentage);

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
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
        [SupportByVersion("Office", 12, 14, 15, 16)]
        void ApplyDataLabels(object type, object iMsoLegendKey, object autoText, object hasLeaderLines, object showSeriesName, object showCategoryName, object showValue, object showPercentage, object showBubbleSize);

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// </summary>
        /// <param name="chartType">NetOffice.OfficeApi.Enums.XlChartType chartType</param>
        /// <param name="typeName">optional object typeName</param>
        [SupportByVersion("Office", 12, 14, 15, 16)]
        void ApplyCustomType(NetOffice.OfficeApi.Enums.XlChartType chartType, object typeName);

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// </summary>
        /// <param name="chartType">NetOffice.OfficeApi.Enums.XlChartType chartType</param>
        [CustomMethod]
        [SupportByVersion("Office", 12, 14, 15, 16)]
        void ApplyCustomType(NetOffice.OfficeApi.Enums.XlChartType chartType);

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// </summary>
        /// <param name="x">Int32 x</param>
        /// <param name="y">Int32 y</param>
        /// <param name="elementID">Int32 elementID</param>
        /// <param name="arg1">Int32 arg1</param>
        /// <param name="arg2">Int32 arg2</param>
        [SupportByVersion("Office", 12, 14, 15, 16)]
        void GetChartElement(Int32 x, Int32 y, Int32 elementID, Int32 arg1, Int32 arg2);

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// </summary>
        /// <param name="source">string source</param>
        /// <param name="plotBy">optional object plotBy</param>
        [SupportByVersion("Office", 12, 14, 15, 16)]
        void SetSourceData(string source, object plotBy);

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// </summary>
        /// <param name="source">string source</param>
        [CustomMethod]
        [SupportByVersion("Office", 12, 14, 15, 16)]
        void SetSourceData(string source);

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// </summary>
        /// <param name="type">optional object type</param>
        /// <param name="axisGroup">optional NetOffice.OfficeApi.Enums.XlAxisGroup AxisGroup = 1</param>
        [SupportByVersion("Office", 12, 14, 15, 16)]
        object Axes(object type, object axisGroup);

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// </summary>
        [CustomMethod]
        [SupportByVersion("Office", 12, 14, 15, 16)]
        object Axes();

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// </summary>
        /// <param name="type">optional object type</param>
        [CustomMethod]
        [SupportByVersion("Office", 12, 14, 15, 16)]
        object Axes(object type);

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// </summary>
        /// <param name="rGallery">Int32 rGallery</param>
        /// <param name="varFormat">optional object varFormat</param>
        [SupportByVersion("Office", 12, 14, 15, 16)]
        void AutoFormat(Int32 rGallery, object varFormat);

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// </summary>
        /// <param name="rGallery">Int32 rGallery</param>
        [CustomMethod]
        [SupportByVersion("Office", 12, 14, 15, 16)]
        void AutoFormat(Int32 rGallery);

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// </summary>
        /// <param name="bstr">string bstr</param>
        [SupportByVersion("Office", 12, 14, 15, 16)]
        void SetBackgroundPicture(string bstr);

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
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
        [SupportByVersion("Office", 12, 14, 15, 16)]
        void ChartWizard(object varSource, object varGallery, object varFormat, object varPlotBy, object varCategoryLabels, object varSeriesLabels, object varHasLegend, object varTitle, object varCategoryTitle, object varValueTitle, object varExtraTitle);

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// </summary>
        [CustomMethod]
        [SupportByVersion("Office", 12, 14, 15, 16)]
        void ChartWizard();

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// </summary>
        /// <param name="varSource">optional object varSource</param>
        [CustomMethod]
        [SupportByVersion("Office", 12, 14, 15, 16)]
        void ChartWizard(object varSource);

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// </summary>
        /// <param name="varSource">optional object varSource</param>
        /// <param name="varGallery">optional object varGallery</param>
        [CustomMethod]
        [SupportByVersion("Office", 12, 14, 15, 16)]
        void ChartWizard(object varSource, object varGallery);

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// </summary>
        /// <param name="varSource">optional object varSource</param>
        /// <param name="varGallery">optional object varGallery</param>
        /// <param name="varFormat">optional object varFormat</param>
        [CustomMethod]
        [SupportByVersion("Office", 12, 14, 15, 16)]
        void ChartWizard(object varSource, object varGallery, object varFormat);

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// </summary>
        /// <param name="varSource">optional object varSource</param>
        /// <param name="varGallery">optional object varGallery</param>
        /// <param name="varFormat">optional object varFormat</param>
        /// <param name="varPlotBy">optional object varPlotBy</param>
        [CustomMethod]
        [SupportByVersion("Office", 12, 14, 15, 16)]
        void ChartWizard(object varSource, object varGallery, object varFormat, object varPlotBy);

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// </summary>
        /// <param name="varSource">optional object varSource</param>
        /// <param name="varGallery">optional object varGallery</param>
        /// <param name="varFormat">optional object varFormat</param>
        /// <param name="varPlotBy">optional object varPlotBy</param>
        /// <param name="varCategoryLabels">optional object varCategoryLabels</param>
        [CustomMethod]
        [SupportByVersion("Office", 12, 14, 15, 16)]
        void ChartWizard(object varSource, object varGallery, object varFormat, object varPlotBy, object varCategoryLabels);

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// </summary>
        /// <param name="varSource">optional object varSource</param>
        /// <param name="varGallery">optional object varGallery</param>
        /// <param name="varFormat">optional object varFormat</param>
        /// <param name="varPlotBy">optional object varPlotBy</param>
        /// <param name="varCategoryLabels">optional object varCategoryLabels</param>
        /// <param name="varSeriesLabels">optional object varSeriesLabels</param>
        [CustomMethod]
        [SupportByVersion("Office", 12, 14, 15, 16)]
        void ChartWizard(object varSource, object varGallery, object varFormat, object varPlotBy, object varCategoryLabels, object varSeriesLabels);

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// </summary>
        /// <param name="varSource">optional object varSource</param>
        /// <param name="varGallery">optional object varGallery</param>
        /// <param name="varFormat">optional object varFormat</param>
        /// <param name="varPlotBy">optional object varPlotBy</param>
        /// <param name="varCategoryLabels">optional object varCategoryLabels</param>
        /// <param name="varSeriesLabels">optional object varSeriesLabels</param>
        /// <param name="varHasLegend">optional object varHasLegend</param>
        [CustomMethod]
        [SupportByVersion("Office", 12, 14, 15, 16)]
        void ChartWizard(object varSource, object varGallery, object varFormat, object varPlotBy, object varCategoryLabels, object varSeriesLabels, object varHasLegend);

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
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
        [SupportByVersion("Office", 12, 14, 15, 16)]
        void ChartWizard(object varSource, object varGallery, object varFormat, object varPlotBy, object varCategoryLabels, object varSeriesLabels, object varHasLegend, object varTitle);

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
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
        [SupportByVersion("Office", 12, 14, 15, 16)]
        void ChartWizard(object varSource, object varGallery, object varFormat, object varPlotBy, object varCategoryLabels, object varSeriesLabels, object varHasLegend, object varTitle, object varCategoryTitle);

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
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
        [SupportByVersion("Office", 12, 14, 15, 16)]
        void ChartWizard(object varSource, object varGallery, object varFormat, object varPlotBy, object varCategoryLabels, object varSeriesLabels, object varHasLegend, object varTitle, object varCategoryTitle, object varValueTitle);

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// </summary>
        /// <param name="appearance">optional Int32 Appearance = 1</param>
        /// <param name="format">optional Int32 Format = -4147</param>
        /// <param name="size">optional Int32 Size = 2</param>
        [SupportByVersion("Office", 12, 14, 15, 16)]
        void CopyPicture(object appearance, object format, object size);

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// </summary>
        [CustomMethod]
        [SupportByVersion("Office", 12, 14, 15, 16)]
        void CopyPicture();

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// </summary>
        /// <param name="appearance">optional Int32 Appearance = 1</param>
        [CustomMethod]
        [SupportByVersion("Office", 12, 14, 15, 16)]
        void CopyPicture(object appearance);

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// </summary>
        /// <param name="appearance">optional Int32 Appearance = 1</param>
        /// <param name="format">optional Int32 Format = -4147</param>
        [CustomMethod]
        [SupportByVersion("Office", 12, 14, 15, 16)]
        void CopyPicture(object appearance, object format);

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// </summary>
        /// <param name="varName">object varName</param>
        /// <param name="localeID">Int32 localeID</param>
        /// <param name="objType">Int32 objType</param>
        [SupportByVersion("Office", 12, 14, 15, 16)]
        object Evaluate(object varName, Int32 localeID, out Int32 objType);

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// </summary>
        /// <param name="varName">object varName</param>
        /// <param name="localeID">Int32 localeID</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [SupportByVersion("Office", 12, 14, 15, 16)]
        object _Evaluate(object varName, Int32 localeID);

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// </summary>
        /// <param name="varType">optional object varType</param>
        [SupportByVersion("Office", 12, 14, 15, 16)]
        void Paste(object varType);

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// </summary>
        [CustomMethod]
        [SupportByVersion("Office", 12, 14, 15, 16)]
        void Paste();

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// </summary>
        /// <param name="bstr">string bstr</param>
        /// <param name="varFilterName">optional object varFilterName</param>
        /// <param name="varInteractive">optional object varInteractive</param>
        [SupportByVersion("Office", 12, 14, 15, 16)]
        bool Export(string bstr, object varFilterName, object varInteractive);

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// </summary>
        /// <param name="bstr">string bstr</param>
        [CustomMethod]
        [SupportByVersion("Office", 12, 14, 15, 16)]
        bool Export(string bstr);

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// </summary>
        /// <param name="bstr">string bstr</param>
        /// <param name="varFilterName">optional object varFilterName</param>
        [CustomMethod]
        [SupportByVersion("Office", 12, 14, 15, 16)]
        bool Export(string bstr, object varFilterName);

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// </summary>
        /// <param name="varName">object varName</param>
        [SupportByVersion("Office", 12, 14, 15, 16)]
        void SetDefaultChart(object varName);

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// </summary>
        /// <param name="bstrFileName">string bstrFileName</param>
        [SupportByVersion("Office", 12, 14, 15, 16)]
        void ApplyChartTemplate(string bstrFileName);

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// </summary>
        /// <param name="bstrFileName">string bstrFileName</param>
        [SupportByVersion("Office", 12, 14, 15, 16)]
        void SaveChartTemplate(string bstrFileName);

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// </summary>
        [SupportByVersion("Office", 12, 14, 15, 16)]
        void ClearToMatchStyle();

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// </summary>
        [SupportByVersion("Office", 12, 14, 15, 16)]
        void RefreshPivotTable();

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// </summary>
        /// <param name="layout">Int32 layout</param>
        /// <param name="varChartType">optional object varChartType</param>
        [SupportByVersion("Office", 12, 14, 15, 16)]
        void ApplyLayout(Int32 layout, object varChartType);

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// </summary>
        /// <param name="layout">Int32 layout</param>
        [CustomMethod]
        [SupportByVersion("Office", 12, 14, 15, 16)]
        void ApplyLayout(Int32 layout);

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// </summary>
        [SupportByVersion("Office", 12, 14, 15, 16)]
        void Refresh();

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// </summary>
        /// <param name="rHS">NetOffice.OfficeApi.Enums.MsoChartElementType rHS</param>
        [SupportByVersion("Office", 12, 14, 15, 16)]
        void SetElement(NetOffice.OfficeApi.Enums.MsoChartElementType rHS);

        /// <summary>
        /// SupportByVersion Office 14, 15, 16
        /// </summary>
        /// <param name="index">optional object index</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [SupportByVersion("Office", 14, 15, 16)]
        object AreaGroups(object index);

        /// <summary>
        /// SupportByVersion Office 14, 15, 16
        /// </summary>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [CustomMethod]
        [SupportByVersion("Office", 14, 15, 16)]
        object AreaGroups();

        /// <summary>
        /// SupportByVersion Office 14, 15, 16
        /// </summary>
        /// <param name="index">optional object index</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [SupportByVersion("Office", 14, 15, 16)]
        object BarGroups(object index);

        /// <summary>
        /// SupportByVersion Office 14, 15, 16
        /// </summary>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [CustomMethod]
        [SupportByVersion("Office", 14, 15, 16)]
        object BarGroups();

        /// <summary>
        /// SupportByVersion Office 14, 15, 16
        /// </summary>
        /// <param name="index">optional object index</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [SupportByVersion("Office", 14, 15, 16)]
        object ColumnGroups(object index);

        /// <summary>
        /// SupportByVersion Office 14, 15, 16
        /// </summary>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [CustomMethod]
        [SupportByVersion("Office", 14, 15, 16)]
        object ColumnGroups();

        /// <summary>
        /// SupportByVersion Office 14, 15, 16
        /// </summary>
        /// <param name="index">optional object index</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [SupportByVersion("Office", 14, 15, 16)]
        object LineGroups(object index);

        /// <summary>
        /// SupportByVersion Office 14, 15, 16
        /// </summary>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [CustomMethod]
        [SupportByVersion("Office", 14, 15, 16)]
        object LineGroups();

        /// <summary>
        /// SupportByVersion Office 14, 15, 16
        /// </summary>
        /// <param name="index">optional object index</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [SupportByVersion("Office", 14, 15, 16)]
        object PieGroups(object index);

        /// <summary>
        /// SupportByVersion Office 14, 15, 16
        /// </summary>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [CustomMethod]
        [SupportByVersion("Office", 14, 15, 16)]
        object PieGroups();

        /// <summary>
        /// SupportByVersion Office 14, 15, 16
        /// </summary>
        /// <param name="index">optional object index</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [SupportByVersion("Office", 14, 15, 16)]
        object DoughnutGroups(object index);

        /// <summary>
        /// SupportByVersion Office 14, 15, 16
        /// </summary>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [CustomMethod]
        [SupportByVersion("Office", 14, 15, 16)]
        object DoughnutGroups();

        /// <summary>
        /// SupportByVersion Office 14, 15, 16
        /// </summary>
        /// <param name="index">optional object index</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [SupportByVersion("Office", 14, 15, 16)]
        object RadarGroups(object index);

        /// <summary>
        /// SupportByVersion Office 14, 15, 16
        /// </summary>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [CustomMethod]
        [SupportByVersion("Office", 14, 15, 16)]
        object RadarGroups();

        /// <summary>
        /// SupportByVersion Office 14, 15, 16
        /// </summary>
        /// <param name="index">optional object index</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [SupportByVersion("Office", 14, 15, 16)]
        object XYGroups(object index);

        /// <summary>
        /// SupportByVersion Office 14, 15, 16
        /// </summary>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [CustomMethod]
        [SupportByVersion("Office", 14, 15, 16)]
        object XYGroups();

        /// <summary>
        /// SupportByVersion Office 14, 15, 16
        /// </summary>
        [SupportByVersion("Office", 14, 15, 16)]
        object Delete();

        /// <summary>
        /// SupportByVersion Office 14, 15, 16
        /// </summary>
        [SupportByVersion("Office", 14, 15, 16)]
        object Copy();

        /// <summary>
        /// SupportByVersion Office 14, 15, 16
        /// </summary>
        /// <param name="replace">optional object replace</param>
        [SupportByVersion("Office", 14, 15, 16)]
        object Select(object replace);

        /// <summary>
        /// SupportByVersion Office 14, 15, 16
        /// </summary>
        [CustomMethod]
        [SupportByVersion("Office", 14, 15, 16)]
        object Select();

        /// <summary>
        /// SupportByVersion Office 15,16
        /// </summary>
        /// <param name="index">optional object index</param>
        [SupportByVersion("Office", 15, 16)]
        object FullSeriesCollection(object index);

        /// <summary>
        /// SupportByVersion Office 15,16
        /// </summary>
        [CustomMethod]
        [SupportByVersion("Office", 15, 16)]
        object FullSeriesCollection();

        /// <summary>
        /// SupportByVersion Office 15,16
        /// </summary>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [SupportByVersion("Office", 15, 16)]
        void DeleteHiddenContent();

        /// <summary>
        /// SupportByVersion Office 15,16
        /// </summary>
        [SupportByVersion("Office", 15, 16)]
        void ClearToMatchColorStyle();

        #endregion
    }
}
