using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.OWC10Api
{
	/// <summary>
	/// DispatchInterface ChChartSpace 
	/// SupportByVersion OWC10, 1
	/// </summary>
	[SupportByVersion("OWC10", 1)]
	[EntityType(EntityType.IsDispatchInterface), BaseType]
	[TypeId("F5B39A82-1480-11D3-8549-00C04FAC67D7")]
    [CoClassSource(typeof(NetOffice.OWC10Api.ChartSpace))]
	public interface ChChartSpace : ICOMObject
	{
		#region Properties

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get/Set
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		NetOffice.OWC10Api.Enums.ChartChartLayoutEnum ChartLayout { get; set; }

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get/Set
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		Int32 ChartWrapCount { get; set; }

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get/Set
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		bool EnableEvents { get; set; }

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get/Set
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		bool HasChartSpaceLegend { get; set; }

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		Int32 MajorVersion { get; }

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		string MinorVersion { get; }

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		string BuildNumber { get; }

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get/Set
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		bool ScreenUpdating { get; set; }

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		NetOffice.OWC10Api.ChBorder Border { get; }

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		NetOffice.OWC10Api.ChCharts Charts { get; }

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get/Set
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		NetOffice.MSDATASRCApi.DataSource DataSource { get; set; }

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get/Set
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		string DataMember { get; set; }

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		NetOffice.OWC10Api.Enums.ChartDataSourceTypeEnum DataSourceType { get; }

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get/Set
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		bool HasChartSpaceTitle { get; set; }

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		NetOffice.OWC10Api.ChInterior Interior { get; }

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		NetOffice.OWC10Api.ChLegend ChartSpaceLegend { get; }

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		[SupportByVersion("OWC10", 1), ProxyResult]
		object Selection { get; }

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		NetOffice.OWC10Api.Enums.ChartSelectionsEnum SelectionType { get; }

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get/Set
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		NetOffice.OWC10Api.Enums.ChartSelectionMarksEnum HasSelectionMarks { get; set; }

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get/Set
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		bool DisplayPropertyToolbox { get; set; }

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		NetOffice.OWC10Api.ChTitle ChartSpaceTitle { get; }

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		string Version { get; }

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get/Set
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		string XMLData { get; set; }

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		[SupportByVersion("OWC10", 1), ProxyResult]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		object Constants { get; }

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		bool CanUndo { get; }

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get/Set
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		bool AllowLayoutEvents { get; set; }

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get/Set
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		bool AllowRenderEvents { get; set; }

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get/Set
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		bool AllowPointRenderEvents { get; set; }

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get/Set
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		bool Enabled { get; set; }

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		string RevisionNumber { get; }

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get/Set
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		Double PrintQuality3D { get; set; }

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get/Set
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		bool DisplayScreenTips { get; set; }

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get/Set
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		string ConnectionString { get; set; }

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get/Set
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		string CommandText { get; set; }

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		[SupportByVersion("OWC10", 1), ProxyResult]
		object InternalPivotTable { get; }

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get/Set
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		bool HasSeriesByRows { get; set; }

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get/Set
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		NetOffice.OWC10Api.Enums.ChartPlotAggregatesEnum PlotAllAggregates { get; set; }

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get/Set
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		bool HasMultipleCharts { get; set; }

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get/Set
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		bool DisplayFieldList { get; set; }

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get/Set
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		bool HasPassiveAlerts { get; set; }

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get/Set
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		string DataSourceName { get; set; }

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get/Set
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		bool DisplayFieldButtons { get; set; }

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		object SelectionList { get; }

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get/Set
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		bool HasPlotDetails { get; set; }

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get/Set
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		bool AllowScreenTipEvents { get; set; }

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		NetOffice.OWC10Api.OCCommands Commands { get; }

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get/Set
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		bool AllowPropertyToolbox { get; set; }

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get/Set
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		bool AllowGrouping { get; set; }

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get/Set
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		bool AllowFiltering { get; set; }

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		Int32 Top { get; }

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		Int32 Left { get; }

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		Int32 Bottom { get; }

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		Int32 Right { get; }

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get/Set
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		bool HasUnifiedScales { get; set; }

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get/Set
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		bool DisplayToolbar { get; set; }

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		[BaseResult]
		NetOffice.MSComctlLibApi.IToolbar Toolbar { get; }

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		bool ViewOnlyMode { get; }

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get/Set
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		bool IsDirty { get; set; }

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get
		/// </summary>
		/// <param name="index">object index</param>
		[SupportByVersion("OWC10", 1)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		object get_International(object index);

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Alias for get_International
		/// </summary>
		/// <param name="index">object index</param>
		[SupportByVersion("OWC10", 1), Redirect("get_International")]
		object International(object index);

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		NetOffice.OWC10Api.OWCLanguageSettings LanguageSettings { get; }

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		[SupportByVersion("OWC10", 1), ProxyResult]
		object Parent { get; }

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get/Set
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		bool HasRuntimeSelection { get; set; }

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get/Set
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		bool DisplayBranding { get; set; }

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get/Set
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		bool DisplayOfficeLogo { get; set; }

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		NetOffice.OWC10Api.Enums.ChartSelectionsEnum ObjectType { get; }

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("OWC10", 1)]
		void BuildLitChart();

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		/// <param name="filename">string filename</param>
		[SupportByVersion("OWC10", 1)]
		void Load(string filename);

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		void Clear();

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		/// <param name="iTopic">Int32 iTopic</param>
		[SupportByVersion("OWC10", 1)]
		void ShowHelp(Int32 iTopic);

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		/// <param name="filename">optional string Filename = chart.gif</param>
		/// <param name="filterName">optional string FilterName = gif</param>
		/// <param name="width">optional Int32 Width = -1</param>
		/// <param name="height">optional Int32 Height = -1</param>
		[SupportByVersion("OWC10", 1)]
		void ExportPicture(object filename, object filterName, object width, object height);

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		[CustomMethod]
		[SupportByVersion("OWC10", 1)]
		void ExportPicture();

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		/// <param name="filename">optional string Filename = chart.gif</param>
		[CustomMethod]
		[SupportByVersion("OWC10", 1)]
		void ExportPicture(object filename);

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		/// <param name="filename">optional string Filename = chart.gif</param>
		/// <param name="filterName">optional string FilterName = gif</param>
		[CustomMethod]
		[SupportByVersion("OWC10", 1)]
		void ExportPicture(object filename, object filterName);

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		/// <param name="filename">optional string Filename = chart.gif</param>
		/// <param name="filterName">optional string FilterName = gif</param>
		/// <param name="width">optional Int32 Width = -1</param>
		[CustomMethod]
		[SupportByVersion("OWC10", 1)]
		void ExportPicture(object filename, object filterName, object width);

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		void Refresh();

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		void Select();

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		void BeginUndo();

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		void EndUndo();

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		void Undo();

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		/// <param name="x">Int32 x</param>
		/// <param name="y">Int32 y</param>
		[SupportByVersion("OWC10", 1)]
		object RangeFromPoint(Int32 x, Int32 y);

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		/// <param name="dimension">NetOffice.OWC10Api.Enums.ChartDimensionsEnum dimension</param>
		/// <param name="dataSourceIndex">Int32 dataSourceIndex</param>
		/// <param name="dataReference">optional object dataReference</param>
		[SupportByVersion("OWC10", 1)]
		void SetData(NetOffice.OWC10Api.Enums.ChartDimensionsEnum dimension, Int32 dataSourceIndex, object dataReference);

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		/// <param name="dimension">NetOffice.OWC10Api.Enums.ChartDimensionsEnum dimension</param>
		/// <param name="dataSourceIndex">Int32 dataSourceIndex</param>
		[CustomMethod]
		[SupportByVersion("OWC10", 1)]
		void SetData(NetOffice.OWC10Api.Enums.ChartDimensionsEnum dimension, Int32 dataSourceIndex);

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		/// <param name="dz">NetOffice.OWC10Api.Enums.ChartDropZonesEnum dz</param>
		[SupportByVersion("OWC10", 1)]
		NetOffice.OWC10Api.ChDropZone DropZones(NetOffice.OWC10Api.Enums.ChartDropZonesEnum dz);

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		/// <param name="punk">object punk</param>
		/// <param name="lPos">Int32 lPos</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("OWC10", 1)]
		void FieldListAddTo(object punk, Int32 lPos);

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("OWC10", 1)]
		void LocateDataSource();

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		/// <param name="x">Int32 x</param>
		/// <param name="y">Int32 y</param>
		/// <param name="menu">object menu</param>
		[SupportByVersion("OWC10", 1)]
		void ShowContextMenu(Int32 x, Int32 y, object menu);

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		/// <param name="filterName">optional string FilterName = gif</param>
		/// <param name="width">optional Int32 Width = -1</param>
		/// <param name="height">optional Int32 Height = -1</param>
		[SupportByVersion("OWC10", 1)]
		object GetPicture(object filterName, object width, object height);

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		[CustomMethod]
		[SupportByVersion("OWC10", 1)]
		object GetPicture();

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		/// <param name="filterName">optional string FilterName = gif</param>
		[CustomMethod]
		[SupportByVersion("OWC10", 1)]
		object GetPicture(object filterName);

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		/// <param name="filterName">optional string FilterName = gif</param>
		/// <param name="width">optional Int32 Width = -1</param>
		[CustomMethod]
		[SupportByVersion("OWC10", 1)]
		object GetPicture(object filterName, object width);

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		/// <param name="dataReference">string dataReference</param>
		/// <param name="seriesByRows">optional bool SeriesByRows = false</param>
		[SupportByVersion("OWC10", 1)]
		void SetSpreadsheetData(string dataReference, object seriesByRows);

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		/// <param name="dataReference">string dataReference</param>
		[CustomMethod]
		[SupportByVersion("OWC10", 1)]
		void SetSpreadsheetData(string dataReference);

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		void Repaint();

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		/// <param name="dimension">NetOffice.OWC10Api.Enums.ChartDimensionsEnum dimension</param>
		[SupportByVersion("OWC10", 1)]
		Int32 GetDataSourceIndex(NetOffice.OWC10Api.Enums.ChartDimensionsEnum dimension);

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("OWC10", 1)]
		void ClearUndo();

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		void OkToBindToControlByName();

		#endregion
	}
}
