using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.OWC10Api
{
	/// <summary>
	/// DispatchInterface IPivotControl 
	/// SupportByVersion OWC10, 1
	/// </summary>
	[SupportByVersion("OWC10", 1)]
	[EntityType(EntityType.IsDispatchInterface), BaseType]
	[TypeId("F5B39B08-1480-11D3-8549-00C04FAC67D7")]
    [CoClassSource(typeof(NetOffice.OWC10Api.PivotTable))]
    public interface IPivotControl : ICOMObject
	{
		#region Properties

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		NetOffice.OWC10Api.PivotView ActiveView { get; }

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get/Set
		/// Unknown COM Proxy
		/// </summary>
		[SupportByVersion("OWC10", 1), ProxyResult]
		object Selection { get; set; }

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
		NetOffice.OWC10Api.PivotData ActiveData { get; }

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		string Version { get; }

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		bool HasDetails { get; }

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
		bool AllowGrouping { get; set; }

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get/Set
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		bool AllowFiltering { get; set; }

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get/Set
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		bool AllowDetails { get; set; }

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
		bool AllowCustomOrdering { get; set; }

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get/Set
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		bool AutoFit { get; set; }

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
		object BackColor { get; set; }

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get/Set
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		bool DisplayExpandIndicator { get; set; }

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get/Set
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		bool RightToLeft { get; set; }

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get/Set
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		Int32 MaxWidth { get; set; }

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get/Set
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		Int32 MaxHeight { get; set; }

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get/Set
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		Int32 Width { get; set; }

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get/Set
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		Int32 Height { get; set; }

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
		bool DisplayPropertyToolbox { get; set; }

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get/Set
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		bool DisplayFieldList { get; set; }

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
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		NetOffice.OWC10Api.Enums.ProviderType ProviderType { get; }

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
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		NetOffice.OWC10Api.Enums.PivotTableMemberExpandEnum MemberExpand { get; set; }

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get/Set
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		NetOffice.ADODBApi.Connection Connection { get; set; }

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
		bool DisplayAlerts { get; set; }

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		object DataMemberStrings { get; }

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get/Set
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		NetOffice.OWC10Api.PivotClassFactory ClassFactory { get; set; }

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
		Int32 Top { get; }

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		Int32 Hwnd { get; }

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get/Set
		/// Unknown COM Proxy
		/// </summary>
		[SupportByVersion("OWC10", 1), ProxyResult]
		object ActiveObject { get; set; }

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		NetOffice.OWC10Api.OCCommands Commands { get; }

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		bool UserMode { get; }

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get/Set
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		string DataMemberCaption { get; set; }

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get/Set
		/// Unknown COM Proxy
		/// </summary>
		[SupportByVersion("OWC10", 1), ProxyResult]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		object DataSourceEx { get; set; }

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get/Set
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		bool IsDirty { get; set; }

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get/Set
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		string CubeProvider { get; set; }

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		string SelectionType { get; }

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get/Set
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		bool DisplayScreenTips { get; set; }

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
		bool DisplayDesignTimeUI { get; set; }

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
		NetOffice.OWC10Api.Enums.PivotEditModeEnum EditMode { get; }

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		string HTMLData { get; }

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
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		bool DisplayBranding { get; set; }

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get/Set
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		bool DisplayOfficeLogo { get; set; }

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		/// <param name="filename">optional string Filename = </param>
		/// <param name="action">optional NetOffice.OWC10Api.Enums.PivotExportActionEnum Action = 1</param>
		[SupportByVersion("OWC10", 1)]
		void Export(object filename, object action);

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
		[SupportByVersion("OWC10", 1)]
		void Refresh();

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		/// <param name="filename">optional string Filename = pivot.gif</param>
		/// <param name="filterName">optional string FilterName = gif</param>
		/// <param name="width">optional Int32 Width = 1024</param>
		/// <param name="height">optional Int32 Height = 1024</param>
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
		/// <param name="filename">optional string Filename = pivot.gif</param>
		[CustomMethod]
		[SupportByVersion("OWC10", 1)]
		void ExportPicture(object filename);

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		/// <param name="filename">optional string Filename = pivot.gif</param>
		/// <param name="filterName">optional string FilterName = gif</param>
		[CustomMethod]
		[SupportByVersion("OWC10", 1)]
		void ExportPicture(object filename, object filterName);

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		/// <param name="filename">optional string Filename = pivot.gif</param>
		/// <param name="filterName">optional string FilterName = gif</param>
		/// <param name="width">optional Int32 Width = 1024</param>
		[CustomMethod]
		[SupportByVersion("OWC10", 1)]
		void ExportPicture(object filename, object filterName, object width);

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("OWC10", 1)]
		void LocateDataSource();

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		/// <param name="selection">optional object Selection = null (Nothing in visual basic)</param>
		[SupportByVersion("OWC10", 1)]
		void Copy(object selection);

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		[CustomMethod]
		[SupportByVersion("OWC10", 1)]
		void Copy();

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		/// <param name="source">NetOffice.OWC10Api.DropSource source</param>
		/// <param name="dragItem">object dragItem</param>
		/// <param name="target">NetOffice.OWC10Api.DropTarget target</param>
		/// <param name="dwLegalEffect">Int32 dwLegalEffect</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("OWC10", 1)]
		void DoDragDrop(NetOffice.OWC10Api.DropSource source, object dragItem, NetOffice.OWC10Api.DropTarget target, Int32 dwLegalEffect);

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		/// <param name="selection">object selection</param>
		/// <param name="activeObject">object activeObject</param>
		/// <param name="scrollType">optional NetOffice.OWC10Api.Enums.PivotScrollTypeEnum ScrollType = 0</param>
		/// <param name="update">optional bool Update = true</param>
		/// <param name="notify">optional bool Notify = true</param>
		[SupportByVersion("OWC10", 1)]
		void Select(object selection, object activeObject, object scrollType, object update, object notify);

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		/// <param name="selection">object selection</param>
		/// <param name="activeObject">object activeObject</param>
		[CustomMethod]
		[SupportByVersion("OWC10", 1)]
		void Select(object selection, object activeObject);

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		/// <param name="selection">object selection</param>
		/// <param name="activeObject">object activeObject</param>
		/// <param name="scrollType">optional NetOffice.OWC10Api.Enums.PivotScrollTypeEnum ScrollType = 0</param>
		[CustomMethod]
		[SupportByVersion("OWC10", 1)]
		void Select(object selection, object activeObject, object scrollType);

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		/// <param name="selection">object selection</param>
		/// <param name="activeObject">object activeObject</param>
		/// <param name="scrollType">optional NetOffice.OWC10Api.Enums.PivotScrollTypeEnum ScrollType = 0</param>
		/// <param name="update">optional bool Update = true</param>
		[CustomMethod]
		[SupportByVersion("OWC10", 1)]
		void Select(object selection, object activeObject, object scrollType, object update);

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
		/// <param name="initialValue">optional object initialValue</param>
		/// <param name="arrowMode">optional NetOffice.OWC10Api.Enums.PivotArrowModeEnum ArrowMode = 0</param>
		/// <param name="caretPosition">optional NetOffice.OWC10Api.Enums.PivotCaretPositionEnum CaretPosition = 0</param>
		[SupportByVersion("OWC10", 1)]
		void StartEdit(object initialValue, object arrowMode, object caretPosition);

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		[CustomMethod]
		[SupportByVersion("OWC10", 1)]
		void StartEdit();

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		/// <param name="initialValue">optional object initialValue</param>
		[CustomMethod]
		[SupportByVersion("OWC10", 1)]
		void StartEdit(object initialValue);

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		/// <param name="initialValue">optional object initialValue</param>
		/// <param name="arrowMode">optional NetOffice.OWC10Api.Enums.PivotArrowModeEnum ArrowMode = 0</param>
		[CustomMethod]
		[SupportByVersion("OWC10", 1)]
		void StartEdit(object initialValue, object arrowMode);

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		/// <param name="accept">optional bool Accept = true</param>
		[SupportByVersion("OWC10", 1)]
		void EndEdit(object accept);

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		[CustomMethod]
		[SupportByVersion("OWC10", 1)]
		void EndEdit();

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("OWC10", 1)]
		void CancelDragDrop();

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		void OkToBindToControlByName();

		#endregion
	}
}
