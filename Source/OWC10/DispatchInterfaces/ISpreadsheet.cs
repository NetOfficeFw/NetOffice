using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.OWC10Api
{
	/// <summary>
	/// DispatchInterface ISpreadsheet 
	/// SupportByVersion OWC10, 1
	/// </summary>
	[SupportByVersion("OWC10", 1)]
	[EntityType(EntityType.IsDispatchInterface), BaseType]
	[TypeId("F5B39A79-1480-11D3-8549-00C04FAC67D7")]
    [CoClassSource(typeof(NetOffice.OWC10Api.Spreadsheet))]
    public interface ISpreadsheet : ICOMObject
	{
		#region Properties

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		[BaseResult]
		NetOffice.OWC10Api._Range ActiveCell { get; }

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		NetOffice.OWC10Api.Worksheet ActiveSheet { get; }

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		NetOffice.OWC10Api.Window ActiveWindow { get; }

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		NetOffice.OWC10Api.Workbook ActiveWorkbook { get; }

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get/Set
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		bool AllowPropertyToolbox { get; set; }

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		[BaseResult]
		NetOffice.OWC10Api.ISpreadsheet Application { get; }

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get/Set
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		bool AutoFit { get; set; }

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		Int32 Build { get; }

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
		NetOffice.OWC10Api.Enums.XlCalculation Calculation { get; set; }

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		Int32 CalculationVersion { get; }

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
		string Caption { get; set; }

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		[BaseResult]
		NetOffice.OWC10Api._Range Cells { get; }

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		[BaseResult]
		NetOffice.OWC10Api._Range Columns { get; }

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		NetOffice.OWC10Api.OCCommands Commands { get; }

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
		/// Get/Set
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		string CSVData { get; set; }

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get/Set
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		string CSVURL { get; set; }

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get/Set
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		string DataMember { get; set; }

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
		string DataType { get; set; }

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		bool DesignMode { get; }

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get/Set
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		bool Dirty { get; set; }

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
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		bool DisplayColumnHeadings { get; set; }

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get/Set
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		bool DisplayDesignTimeUI { get; set; }

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get/Set
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		bool DisplayGridlines { get; set; }

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get/Set
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		bool DisplayHorizontalScrollBar { get; set; }

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get/Set
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		bool DisplayOfficeLogo { get; set; }

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get/Set
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		bool DisplayPropertyToolbox { get; set; }

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get/Set
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		bool DisplayRowHeadings { get; set; }

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get/Set
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		bool DisplayTitleBar { get; set; }

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get/Set
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		bool DisplayToolbar { get; set; }

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get/Set
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		bool DisplayVerticalScrollBar { get; set; }

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get/Set
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		bool DisplayWorkbookTabs { get; set; }

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
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		bool EnableUndo { get; set; }

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get/Set
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		string HTMLData { get; set; }

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get/Set
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		string HTMLURL { get; set; }

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		Int32 InstanceID { get; }

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
		/// Get/Set
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		object MaxHeight { get; set; }

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get/Set
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		object MaxWidth { get; set; }

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
		/// Get/Set
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		bool MoveAfterReturn { get; set; }

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get/Set
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		NetOffice.OWC10Api.Enums.XlDirection MoveAfterReturnDirection { get; set; }

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		string Name { get; }

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		NetOffice.OWC10Api.Names Names { get; }

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get
		/// </summary>
		/// <param name="cell1">object cell1</param>
		/// <param name="cell2">optional object cell2</param>
		[SupportByVersion("OWC10", 1)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		NetOffice.OWC10Api._Range get_Range(object cell1, object cell2);

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Alias for get_Range
		/// </summary>
		/// <param name="cell1">object cell1</param>
		/// <param name="cell2">optional object cell2</param>
		[SupportByVersion("OWC10", 1), Redirect("get_Range")]
		NetOffice.OWC10Api._Range Range(object cell1, object cell2);

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get
		/// </summary>
		/// <param name="cell1">object cell1</param>
		[SupportByVersion("OWC10", 1)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		NetOffice.OWC10Api._Range get_Range(object cell1);

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Alias for get_Range
		/// </summary>
		/// <param name="cell1">object cell1</param>
		[SupportByVersion("OWC10", 1), Redirect("get_Range")]
		NetOffice.OWC10Api._Range Range(object cell1);

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
		bool RightToLeft { get; set; }

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		[BaseResult]
		NetOffice.OWC10Api._Range Rows { get; }

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get/Set
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		bool ScreenUpdating { get; set; }

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		[BaseResult]
		NetOffice.OWC10Api._Range Selection { get; }

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		NetOffice.OWC10Api.Sheets Sheets { get; }

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		NetOffice.OWC10Api.TitleBar TitleBar { get; }

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
		string Value { get; }

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
		string ViewableRange { get; set; }

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		bool ViewOnlyMode { get; }

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		NetOffice.OWC10Api.Windows Windows { get; }

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		NetOffice.OWC10Api.Workbooks Workbooks { get; }

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		NetOffice.OWC10Api.Worksheets Worksheets { get; }

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get/Set
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		string XMLData { get; set; }

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get/Set
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		string XMLURL { get; set; }

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		/// <param name="addIn">object addIn</param>
		[SupportByVersion("OWC10", 1)]
		void AddIn(object addIn);

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		void BeginUndo();

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		void Calculate();

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		void CalculateFull();

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		/// <param name="cancel">optional bool Cancel = false</param>
		[SupportByVersion("OWC10", 1)]
		void EndUndo(object cancel);

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		[CustomMethod]
		[SupportByVersion("OWC10", 1)]
		void EndUndo();

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		/// <param name="expression">object expression</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("OWC10", 1)]
		object _Evaluate(object expression);

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		/// <param name="expression">object expression</param>
		[SupportByVersion("OWC10", 1)]
		object Evaluate(object expression);

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		/// <param name="filename">optional string Filename = </param>
		/// <param name="action">optional NetOffice.OWC10Api.Enums.SheetExportActionEnum Action = 1</param>
		/// <param name="format">optional NetOffice.OWC10Api.Enums.SheetExportFormat Format = 0</param>
		[SupportByVersion("OWC10", 1)]
		void Export(object filename, object action, object format);

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		[CustomMethod]
		[SupportByVersion("OWC10", 1)]
		void Export();

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		/// <param name="filename">optional string Filename = </param>
		[CustomMethod]
		[SupportByVersion("OWC10", 1)]
		void Export(object filename);

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		/// <param name="filename">optional string Filename = </param>
		/// <param name="action">optional NetOffice.OWC10Api.Enums.SheetExportActionEnum Action = 1</param>
		[CustomMethod]
		[SupportByVersion("OWC10", 1)]
		void Export(object filename, object action);

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("OWC10", 1)]
		void LocateDataSource();

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		/// <param name="bstr">string bstr</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("OWC10", 1)]
		object msDataSourceObject(string bstr);

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		/// <param name="range1">NetOffice.OWC10Api._Range range1</param>
		/// <param name="range2">NetOffice.OWC10Api._Range range2</param>
		[SupportByVersion("OWC10", 1)]
		[BaseResult]
		NetOffice.OWC10Api._Range RectIntersect(NetOffice.OWC10Api._Range range1, NetOffice.OWC10Api._Range range2);

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		/// <param name="range1">NetOffice.OWC10Api._Range range1</param>
		/// <param name="range2">NetOffice.OWC10Api._Range range2</param>
		[SupportByVersion("OWC10", 1)]
		[BaseResult]
		NetOffice.OWC10Api._Range RectUnion(NetOffice.OWC10Api._Range range1, NetOffice.OWC10Api._Range range2);

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		void Refresh();

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		void Repaint();

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		void ShowAbout();

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
		/// <param name="topic">Int32 topic</param>
		[SupportByVersion("OWC10", 1)]
		void ShowHelp(Int32 topic);

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		void Undo();

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		void UpdatePropertyToolbox();

		#endregion
	}
}
