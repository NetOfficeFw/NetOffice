using System.Reflection;
using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.VisioApi
{
	/// <summary>
	/// DispatchInterface IVDispExtender 
	/// SupportByVersion Visio, 11,12,14,15,16
	/// </summary>
	[SupportByVersion("Visio", 11,12,14,15,16)]
	[EntityType(EntityType.IsDispatchInterface), BaseType]
	[TypeId("000D0D0F-0000-0000-C000-000000000046")]
    [CoClassSource(typeof(NetOffice.VisioApi.Extender))]
    public interface IVDispExtender : ICOMObject
	{
		#region Properties

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		string Name { get; set; }

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16), ProxyResult]
		object Object { get; }

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16), ProxyResult]
		object Parent { get; }

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		[BaseResult]
		NetOffice.VisioApi.IVShape Shape { get; }

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		[BaseResult]
		NetOffice.VisioApi.IVDocument Document { get; }

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16), ProxyResult]
		object ShapeParent { get; }

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		[BaseResult]
		NetOffice.VisioApi.IVApplication Application { get; }

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		[BaseResult]
		NetOffice.VisioApi.IVMaster Master { get; }

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <param name="localeSpecificCellName">string localeSpecificCellName</param>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		NetOffice.VisioApi.IVCell get_Cells(string localeSpecificCellName);

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Alias for get_Cells
		/// </summary>
		/// <param name="localeSpecificCellName">string localeSpecificCellName</param>
		[SupportByVersion("Visio", 11,12,14,15,16), Redirect("get_Cells")]
		NetOffice.VisioApi.IVCell Cells(string localeSpecificCellName);

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <param name="section">Int16 section</param>
		/// <param name="row">Int16 row</param>
		/// <param name="column">Int16 column</param>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		NetOffice.VisioApi.IVCell get_CellsSRC(Int16 section, Int16 row, Int16 column);

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Alias for get_CellsSRC
		/// </summary>
		/// <param name="section">Int16 section</param>
		/// <param name="row">Int16 row</param>
		/// <param name="column">Int16 column</param>
		[SupportByVersion("Visio", 11,12,14,15,16), Redirect("get_CellsSRC")]
		NetOffice.VisioApi.IVCell CellsSRC(Int16 section, Int16 row, Int16 column);

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		string Data1 { get; set; }

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		string Data2 { get; set; }

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		string Data3 { get; set; }

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		string Help { get; set; }

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		string NameID { get; }

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <param name="section">Int16 section</param>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		Int16 get_RowCount(Int16 section);

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Alias for get_RowCount
		/// </summary>
		/// <param name="section">Int16 section</param>
		[SupportByVersion("Visio", 11,12,14,15,16), Redirect("get_RowCount")]
		Int16 RowCount(Int16 section);

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <param name="section">Int16 section</param>
		/// <param name="row">Int16 row</param>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		Int16 get_RowsCellCount(Int16 section, Int16 row);

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Alias for get_RowsCellCount
		/// </summary>
		/// <param name="section">Int16 section</param>
		/// <param name="row">Int16 row</param>
		[SupportByVersion("Visio", 11,12,14,15,16), Redirect("get_RowsCellCount")]
		Int16 RowsCellCount(Int16 section, Int16 row);

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <param name="section">Int16 section</param>
		/// <param name="row">Int16 row</param>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		Int16 get_RowType(Int16 section, Int16 row);

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <param name="section">Int16 section</param>
        /// <param name="row">Int16 row</param>
        /// <param name="value">Int16 value</param>
        [SupportByVersion("Visio", 11,12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		void set_RowType(Int16 section, Int16 row, Int16 value);

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Alias for get_RowType
		/// </summary>
		/// <param name="section">Int16 section</param>
		/// <param name="row">Int16 row</param>
		[SupportByVersion("Visio", 11,12,14,15,16), Redirect("get_RowType")]
		Int16 RowType(Int16 section, Int16 row);

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		[BaseResult]
		NetOffice.VisioApi.IVConnects Connects { get; }

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		Int16 ShapeIndex16 { get; }

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		string Style { get; set; }

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		string StyleKeepFmt { get; set; }

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		string LineStyle { get; set; }

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		string LineStyleKeepFmt { get; set; }

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		string FillStyle { get; set; }

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		string FillStyleKeepFmt { get; set; }

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <param name="fUniqueID">Int16 fUniqueID</param>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		string get_UniqueID(Int16 fUniqueID);

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Alias for get_UniqueID
		/// </summary>
		/// <param name="fUniqueID">Int16 fUniqueID</param>
		[SupportByVersion("Visio", 11,12,14,15,16), Redirect("get_UniqueID")]
		string UniqueID(Int16 fUniqueID);

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		[BaseResult]
		NetOffice.VisioApi.IVPage ContainingPage { get; }

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		[BaseResult]
		NetOffice.VisioApi.IVMaster ContainingMaster { get; }

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		[BaseResult]
		NetOffice.VisioApi.IVShape ContainingShape { get; }

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <param name="section">Int16 section</param>
		/// <param name="fExistsLocally">Int16 fExistsLocally</param>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		Int16 get_SectionExists(Int16 section, Int16 fExistsLocally);

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Alias for get_SectionExists
		/// </summary>
		/// <param name="section">Int16 section</param>
		/// <param name="fExistsLocally">Int16 fExistsLocally</param>
		[SupportByVersion("Visio", 11,12,14,15,16), Redirect("get_SectionExists")]
		Int16 SectionExists(Int16 section, Int16 fExistsLocally);

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <param name="section">Int16 section</param>
		/// <param name="row">Int16 row</param>
		/// <param name="fExistsLocally">Int16 fExistsLocally</param>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		Int16 get_RowExists(Int16 section, Int16 row, Int16 fExistsLocally);

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Alias for get_RowExists
		/// </summary>
		/// <param name="section">Int16 section</param>
		/// <param name="row">Int16 row</param>
		/// <param name="fExistsLocally">Int16 fExistsLocally</param>
		[SupportByVersion("Visio", 11,12,14,15,16), Redirect("get_RowExists")]
		Int16 RowExists(Int16 section, Int16 row, Int16 fExistsLocally);

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <param name="localeSpecificCellName">string localeSpecificCellName</param>
		/// <param name="fExistsLocally">Int16 fExistsLocally</param>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		Int16 get_CellExists(string localeSpecificCellName, Int16 fExistsLocally);

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Alias for get_CellExists
		/// </summary>
		/// <param name="localeSpecificCellName">string localeSpecificCellName</param>
		/// <param name="fExistsLocally">Int16 fExistsLocally</param>
		[SupportByVersion("Visio", 11,12,14,15,16), Redirect("get_CellExists")]
		Int16 CellExists(string localeSpecificCellName, Int16 fExistsLocally);

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <param name="section">Int16 section</param>
		/// <param name="row">Int16 row</param>
		/// <param name="column">Int16 column</param>
		/// <param name="fExistsLocally">Int16 fExistsLocally</param>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		Int16 get_CellsSRCExists(Int16 section, Int16 row, Int16 column, Int16 fExistsLocally);

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Alias for get_CellsSRCExists
		/// </summary>
		/// <param name="section">Int16 section</param>
		/// <param name="row">Int16 row</param>
		/// <param name="column">Int16 column</param>
		/// <param name="fExistsLocally">Int16 fExistsLocally</param>
		[SupportByVersion("Visio", 11,12,14,15,16), Redirect("get_CellsSRCExists")]
		Int16 CellsSRCExists(Int16 section, Int16 row, Int16 column, Int16 fExistsLocally);

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		Int16 LayerCount { get; }

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <param name="index">Int16 index</param>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		NetOffice.VisioApi.IVLayer get_Layer(Int16 index);

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Alias for get_Layer
		/// </summary>
		/// <param name="index">Int16 index</param>
		[SupportByVersion("Visio", 11,12,14,15,16), Redirect("get_Layer")]
		NetOffice.VisioApi.IVLayer Layer(Int16 index);

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		[BaseResult]
		NetOffice.VisioApi.IVEventList EventList { get; }

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		Int16 PersistsEvents { get; }

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		string ClassID { get; }

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16), ProxyResult]
		object ShapeObject { get; }

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		Int16 ShapeID16 { get; }

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		[BaseResult]
		NetOffice.VisioApi.IVConnects FromConnects { get; }

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		[BaseResult]
		NetOffice.VisioApi.IVHyperlink Hyperlink { get; }

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		string ProgID { get; }

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		Int16 ObjectIsInherited { get; }

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		Int32 ShapeID { get; }

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		Int32 ShapeIndex { get; }

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// </summary>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("Visio", 11,12,14,15,16)]
		void Delete();

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// </summary>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("Visio", 11,12,14,15,16)]
		void Index();

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// </summary>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("Visio", 11,12,14,15,16)]
		void VoidGroup();

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		void BringForward();

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		void BringToFront();

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		void ConvertToGroup();

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		void SendBackward();

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		void SendToBack();

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		void ShapeCopy();

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		void ShapeCut();

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		void ShapeDelete();

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// </summary>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("Visio", 11,12,14,15,16)]
		void VoidShapeDuplicate();

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="section">Int16 section</param>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		Int16 AddSection(Int16 section);

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="section">Int16 section</param>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		void DeleteSection(Int16 section);

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="section">Int16 section</param>
		/// <param name="row">Int16 row</param>
		/// <param name="rowTag">Int16 rowTag</param>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		Int16 AddRow(Int16 section, Int16 row, Int16 rowTag);

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="section">Int16 section</param>
		/// <param name="row">Int16 row</param>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		void DeleteRow(Int16 section, Int16 row);

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="xPos">Double xPos</param>
		/// <param name="yPos">Double yPos</param>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		void SetCenter(Double xPos, Double yPos);

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="fileName">string fileName</param>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		void Export(string fileName);

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="section">Int16 section</param>
		/// <param name="rowName">string rowName</param>
		/// <param name="rowTag">Int16 rowTag</param>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		Int16 AddNamedRow(Int16 section, string rowName, Int16 rowTag);

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="section">Int16 section</param>
		/// <param name="row">Int16 row</param>
		/// <param name="rowTag">Int16 rowTag</param>
		/// <param name="rowCount">Int16 rowCount</param>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		Int16 AddRows(Int16 section, Int16 row, Int16 rowTag, Int16 rowCount);

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		[BaseResult]
		NetOffice.VisioApi.IVWindow OpenSheetWindow();

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="sRCStream">Int16[] sRCStream</param>
		/// <param name="formulaArray">object[] formulaArray</param>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		void GetFormulas(Int16[] sRCStream, out object[] formulaArray);

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="sRCStream">Int16[] sRCStream</param>
		/// <param name="flags">Int16 flags</param>
		/// <param name="unitsNamesOrCodes">object[] unitsNamesOrCodes</param>
		/// <param name="resultArray">object[] resultArray</param>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		void GetResults(Int16[] sRCStream, Int16 flags, object[] unitsNamesOrCodes, out object[] resultArray);

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="sRCStream">Int16[] sRCStream</param>
		/// <param name="formulaArray">object[] formulaArray</param>
		/// <param name="flags">Int16 flags</param>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		Int16 SetFormulas(Int16[] sRCStream, object[] formulaArray, Int16 flags);

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="sRCStream">Int16[] sRCStream</param>
		/// <param name="unitsNamesOrCodes">object[] unitsNamesOrCodes</param>
		/// <param name="resultArray">object[] resultArray</param>
		/// <param name="flags">Int16 flags</param>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		Int16 SetResults(Int16[] sRCStream, object[] unitsNamesOrCodes, object[] resultArray, Int16 flags);

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="flags">Int16 flags</param>
		/// <param name="lpr8Left">Double lpr8Left</param>
		/// <param name="lpr8Bottom">Double lpr8Bottom</param>
		/// <param name="lpr8Right">Double lpr8Right</param>
		/// <param name="lpr8Top">Double lpr8Top</param>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		void BoundingBox(Int16 flags, out Double lpr8Left, out Double lpr8Bottom, out Double lpr8Right, out Double lpr8Top);

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="xPos">Double xPos</param>
		/// <param name="yPos">Double yPos</param>
		/// <param name="tolerance">Double tolerance</param>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		Int16 HitTest(Double xPos, Double yPos, Double tolerance);

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		[BaseResult]
		NetOffice.VisioApi.IVShape Group();

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		[BaseResult]
		NetOffice.VisioApi.IVShape ShapeDuplicate();

		#endregion
	}
}
