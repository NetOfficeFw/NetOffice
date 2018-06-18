using System.Reflection;
using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.VisioApi
{
    /// <summary>
    /// LPVISIOSHAPE
    /// </summary>
    [SyntaxBypass]
    public interface LPVISIOSHAPE_ : ICOMObject
    {
        #region Properties

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <param name="fIncludeSubShapes">optional bool fIncludeSubShapes</param>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        Double get_AreaIU(object fIncludeSubShapes);

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// Alias for get_AreaIU
        /// </summary>
        /// <param name="fIncludeSubShapes">optional bool fIncludeSubShapes</param>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16), Redirect("get_AreaIU")]
        Double AreaIU(object fIncludeSubShapes);

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <param name="fIncludeSubShapes">optional bool fIncludeSubShapes</param>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        Double get_LengthIU(object fIncludeSubShapes);

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// Alias for get_LengthIU
        /// </summary>
        /// <param name="fIncludeSubShapes">optional bool fIncludeSubShapes</param>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16), Redirect("get_LengthIU")]
        Double LengthIU(object fIncludeSubShapes);

        #endregion
    }

    /// <summary>
    /// Interface LPVISIOSHAPE 
    /// SupportByVersion Visio, 11,12,14,15,16
    /// </summary>
    [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
    [EntityType(EntityType.IsInterface)]
	[TypeId("00000000-0000-0000-0000-000000000000")]
    public interface LPVISIOSHAPE : LPVISIOSHAPE_
    {
        #region Properties

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        [BaseResult]
        NetOffice.VisioApi.IVDocument Document { get; }

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// Get/Set
        /// Unknown COM Proxy
        /// </summary>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16), ProxyResult]
        object Parent { get; set; }

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        [BaseResult]
        NetOffice.VisioApi.IVApplication Application { get; }

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        Int16 Stat { get; }

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        [BaseResult]
        NetOffice.VisioApi.IVMaster Master { get; }

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        Int16 Type { get; }

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        Int16 ObjectType { get; }

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <param name="localeSpecificCellName">string localeSpecificCellName</param>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        NetOffice.VisioApi.IVCell get_Cells(string localeSpecificCellName);

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// Alias for get_Cells
        /// </summary>
        /// <param name="localeSpecificCellName">string localeSpecificCellName</param>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16), Redirect("get_Cells")]
        NetOffice.VisioApi.IVCell Cells(string localeSpecificCellName);

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <param name="section">Int16 section</param>
        /// <param name="row">Int16 row</param>
        /// <param name="column">Int16 column</param>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        NetOffice.VisioApi.IVCell get_CellsSRC(Int16 section, Int16 row, Int16 column);

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// Alias for get_CellsSRC
        /// </summary>
        /// <param name="section">Int16 section</param>
        /// <param name="row">Int16 row</param>
        /// <param name="column">Int16 column</param>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16), Redirect("get_CellsSRC")]
        NetOffice.VisioApi.IVCell CellsSRC(Int16 section, Int16 row, Int16 column);

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        [BaseResult]
        NetOffice.VisioApi.IVShapes Shapes { get; }

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        string Data1 { get; set; }

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        string Data2 { get; set; }

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        string Data3 { get; set; }

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        string Help { get; set; }

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        string NameID { get; }

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        string Name { get; set; }

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        string Text { get; set; }

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        Int32 CharCount { get; }

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        [BaseResult]
        NetOffice.VisioApi.IVCharacters Characters { get; }

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        Int16 OneD { get; set; }

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        Int16 GeometryCount { get; }

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <param name="section">Int16 section</param>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        Int16 get_RowCount(Int16 section);

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// Alias for get_RowCount
        /// </summary>
        /// <param name="section">Int16 section</param>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16), Redirect("get_RowCount")]
        Int16 RowCount(Int16 section);

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <param name="section">Int16 section</param>
        /// <param name="row">Int16 row</param>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        Int16 get_RowsCellCount(Int16 section, Int16 row);

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// Alias for get_RowsCellCount
        /// </summary>
        /// <param name="section">Int16 section</param>
        /// <param name="row">Int16 row</param>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16), Redirect("get_RowsCellCount")]
        Int16 RowsCellCount(Int16 section, Int16 row);

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <param name="section">Int16 section</param>
        /// <param name="row">Int16 row</param>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        Int16 get_RowType(Int16 section, Int16 row);

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <param name="section">Int16 section</param>
        /// <param name="row">Int16 row</param>
        /// <param name="value">Int16 value</param>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        void set_RowType(Int16 section, Int16 row, Int16 value);

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// Alias for get_RowType
        /// </summary>
        /// <param name="section">Int16 section</param>
        /// <param name="row">Int16 row</param>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16), Redirect("get_RowType")]
        Int16 RowType(Int16 section, Int16 row);

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        [BaseResult]
        NetOffice.VisioApi.IVConnects Connects { get; }

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        Int16 Index16 { get; }

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        string Style { get; set; }

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        string StyleKeepFmt { get; set; }

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        string LineStyle { get; set; }

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        string LineStyleKeepFmt { get; set; }

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        string FillStyle { get; set; }

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        string FillStyleKeepFmt { get; set; }

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        string TextStyle { get; set; }

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        string TextStyleKeepFmt { get; set; }

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        Double old_AreaIU { get; }

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        Double old_LengthIU { get; }

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// Get
        /// Unknown COM Proxy
        /// </summary>
        /// <param name="fFill">Int16 fFill</param>
        /// <param name="lineRes">Double lineRes</param>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16), ProxyResult]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        object get_GeomExIf(Int16 fFill, Double lineRes);

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// Alias for get_GeomExIf
        /// Unknown COM Proxy
        /// </summary>
        /// <param name="fFill">Int16 fFill</param>
        /// <param name="lineRes">Double lineRes</param>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16), ProxyResult, Redirect("get_GeomExIf")]
        object GeomExIf(Int16 fFill, Double lineRes);

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <param name="fUniqueID">Int16 fUniqueID</param>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        string get_UniqueID(Int16 fUniqueID);

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// Alias for get_UniqueID
        /// </summary>
        /// <param name="fUniqueID">Int16 fUniqueID</param>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16), Redirect("get_UniqueID")]
        string UniqueID(Int16 fUniqueID);

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        [BaseResult]
        NetOffice.VisioApi.IVPage ContainingPage { get; }

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        [BaseResult]
        NetOffice.VisioApi.IVMaster ContainingMaster { get; }

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        [BaseResult]
        NetOffice.VisioApi.IVShape ContainingShape { get; }

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <param name="section">Int16 section</param>
        /// <param name="fExistsLocally">Int16 fExistsLocally</param>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        Int16 get_SectionExists(Int16 section, Int16 fExistsLocally);

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// Alias for get_SectionExists
        /// </summary>
        /// <param name="section">Int16 section</param>
        /// <param name="fExistsLocally">Int16 fExistsLocally</param>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16), Redirect("get_SectionExists")]
        Int16 SectionExists(Int16 section, Int16 fExistsLocally);

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <param name="section">Int16 section</param>
        /// <param name="row">Int16 row</param>
        /// <param name="fExistsLocally">Int16 fExistsLocally</param>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        Int16 get_RowExists(Int16 section, Int16 row, Int16 fExistsLocally);

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// Alias for get_RowExists
        /// </summary>
        /// <param name="section">Int16 section</param>
        /// <param name="row">Int16 row</param>
        /// <param name="fExistsLocally">Int16 fExistsLocally</param>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16), Redirect("get_RowExists")]
        Int16 RowExists(Int16 section, Int16 row, Int16 fExistsLocally);

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <param name="localeSpecificCellName">string localeSpecificCellName</param>
        /// <param name="fExistsLocally">Int16 fExistsLocally</param>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        Int16 get_CellExists(string localeSpecificCellName, Int16 fExistsLocally);

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// Alias for get_CellExists
        /// </summary>
        /// <param name="localeSpecificCellName">string localeSpecificCellName</param>
        /// <param name="fExistsLocally">Int16 fExistsLocally</param>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16), Redirect("get_CellExists")]
        Int16 CellExists(string localeSpecificCellName, Int16 fExistsLocally);

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <param name="section">Int16 section</param>
        /// <param name="row">Int16 row</param>
        /// <param name="column">Int16 column</param>
        /// <param name="fExistsLocally">Int16 fExistsLocally</param>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
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
        [SupportByVersion("Visio", 11, 12, 14, 15, 16), Redirect("get_CellsSRCExists")]
        Int16 CellsSRCExists(Int16 section, Int16 row, Int16 column, Int16 fExistsLocally);

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        Int16 LayerCount { get; }

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <param name="index">Int16 index</param>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        NetOffice.VisioApi.IVLayer get_Layer(Int16 index);

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// Alias for get_Layer
        /// </summary>
        /// <param name="index">Int16 index</param>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16), Redirect("get_Layer")]
        NetOffice.VisioApi.IVLayer Layer(Int16 index);

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        [BaseResult]
        NetOffice.VisioApi.IVEventList EventList { get; }

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        Int16 PersistsEvents { get; }

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        string ClassID { get; }

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        Int16 ForeignType { get; }

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// Get
        /// Unknown COM Proxy
        /// </summary>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16), ProxyResult]
        object Object { get; }

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        Int16 ID16 { get; }

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        [BaseResult]
        NetOffice.VisioApi.IVConnects FromConnects { get; }

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        [BaseResult]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        NetOffice.VisioApi.IVHyperlink Hyperlink { get; }

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        string ProgID { get; }

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        Int16 ObjectIsInherited { get; }

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        [BaseResult]
        NetOffice.VisioApi.IVPaths Paths { get; }

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        [BaseResult]
        NetOffice.VisioApi.IVPaths PathsLocal { get; }

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        Int32 ID { get; }

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        Int32 Index { get; }

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <param name="index">Int16 index</param>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        NetOffice.VisioApi.IVSection get_Section(Int16 index);

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// Alias for get_Section
        /// </summary>
        /// <param name="index">Int16 index</param>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16), Redirect("get_Section")]
        NetOffice.VisioApi.IVSection Section(Int16 index);

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        [BaseResult]
        NetOffice.VisioApi.IVHyperlinks Hyperlinks { get; }

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <param name="otherShape">NetOffice.VisioApi.IVShape otherShape</param>
        /// <param name="tolerance">Double tolerance</param>
        /// <param name="flags">Int16 flags</param>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        Int16 get_SpatialRelation(NetOffice.VisioApi.IVShape otherShape, Double tolerance, Int16 flags);

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// Alias for get_SpatialRelation
        /// </summary>
        /// <param name="otherShape">NetOffice.VisioApi.IVShape otherShape</param>
        /// <param name="tolerance">Double tolerance</param>
        /// <param name="flags">Int16 flags</param>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16), Redirect("get_SpatialRelation")]
        Int16 SpatialRelation(NetOffice.VisioApi.IVShape otherShape, Double tolerance, Int16 flags);

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <param name="otherShape">NetOffice.VisioApi.IVShape otherShape</param>
        /// <param name="flags">Int16 flags</param>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        Double get_DistanceFrom(NetOffice.VisioApi.IVShape otherShape, Int16 flags);

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// Alias for get_DistanceFrom
        /// </summary>
        /// <param name="otherShape">NetOffice.VisioApi.IVShape otherShape</param>
        /// <param name="flags">Int16 flags</param>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16), Redirect("get_DistanceFrom")]
        Double DistanceFrom(NetOffice.VisioApi.IVShape otherShape, Int16 flags);

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <param name="x">Double x</param>
        /// <param name="y">Double y</param>
        /// <param name="flags">Int16 flags</param>
        /// <param name="pvPathIndex">optional object pvPathIndex</param>
        /// <param name="pvCurveIndex">optional object pvCurveIndex</param>
        /// <param name="pvt">optional object pvt</param>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        Double get_DistanceFromPoint(Double x, Double y, Int16 flags, object pvPathIndex, object pvCurveIndex, object pvt);

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// Alias for get_DistanceFromPoint
        /// </summary>
        /// <param name="x">Double x</param>
        /// <param name="y">Double y</param>
        /// <param name="flags">Int16 flags</param>
        /// <param name="pvPathIndex">optional object pvPathIndex</param>
        /// <param name="pvCurveIndex">optional object pvCurveIndex</param>
        /// <param name="pvt">optional object pvt</param>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16), Redirect("get_DistanceFromPoint")]
        Double DistanceFromPoint(Double x, Double y, Int16 flags, object pvPathIndex, object pvCurveIndex, object pvt);

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <param name="x">Double x</param>
        /// <param name="y">Double y</param>
        /// <param name="flags">Int16 flags</param>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        Double get_DistanceFromPoint(Double x, Double y, Int16 flags);

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// Alias for get_DistanceFromPoint
        /// </summary>
        /// <param name="x">Double x</param>
        /// <param name="y">Double y</param>
        /// <param name="flags">Int16 flags</param>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16), Redirect("get_DistanceFromPoint")]
        Double DistanceFromPoint(Double x, Double y, Int16 flags);

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <param name="x">Double x</param>
        /// <param name="y">Double y</param>
        /// <param name="flags">Int16 flags</param>
        /// <param name="pvPathIndex">optional object pvPathIndex</param>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        Double get_DistanceFromPoint(Double x, Double y, Int16 flags, object pvPathIndex);

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// Alias for get_DistanceFromPoint
        /// </summary>
        /// <param name="x">Double x</param>
        /// <param name="y">Double y</param>
        /// <param name="flags">Int16 flags</param>
        /// <param name="pvPathIndex">optional object pvPathIndex</param>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16), Redirect("get_DistanceFromPoint")]
        Double DistanceFromPoint(Double x, Double y, Int16 flags, object pvPathIndex);

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <param name="x">Double x</param>
        /// <param name="y">Double y</param>
        /// <param name="flags">Int16 flags</param>
        /// <param name="pvPathIndex">optional object pvPathIndex</param>
        /// <param name="pvCurveIndex">optional object pvCurveIndex</param>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        Double get_DistanceFromPoint(Double x, Double y, Int16 flags, object pvPathIndex, object pvCurveIndex);

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// Alias for get_DistanceFromPoint
        /// </summary>
        /// <param name="x">Double x</param>
        /// <param name="y">Double y</param>
        /// <param name="flags">Int16 flags</param>
        /// <param name="pvPathIndex">optional object pvPathIndex</param>
        /// <param name="pvCurveIndex">optional object pvCurveIndex</param>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16), Redirect("get_DistanceFromPoint")]
        Double DistanceFromPoint(Double x, Double y, Int16 flags, object pvPathIndex, object pvCurveIndex);

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <param name="relation">Int16 relation</param>
        /// <param name="tolerance">Double tolerance</param>
        /// <param name="flags">Int16 flags</param>
        /// <param name="resultRoot">optional object resultRoot</param>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        NetOffice.VisioApi.IVSelection get_SpatialNeighbors(Int16 relation, Double tolerance, Int16 flags, object resultRoot);

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// Alias for get_SpatialNeighbors
        /// </summary>
        /// <param name="relation">Int16 relation</param>
        /// <param name="tolerance">Double tolerance</param>
        /// <param name="flags">Int16 flags</param>
        /// <param name="resultRoot">optional object resultRoot</param>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16), Redirect("get_SpatialNeighbors")]
        NetOffice.VisioApi.IVSelection SpatialNeighbors(Int16 relation, Double tolerance, Int16 flags, object resultRoot);

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <param name="relation">Int16 relation</param>
        /// <param name="tolerance">Double tolerance</param>
        /// <param name="flags">Int16 flags</param>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        NetOffice.VisioApi.IVSelection get_SpatialNeighbors(Int16 relation, Double tolerance, Int16 flags);

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// Alias for get_SpatialNeighbors
        /// </summary>
        /// <param name="relation">Int16 relation</param>
        /// <param name="tolerance">Double tolerance</param>
        /// <param name="flags">Int16 flags</param>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16), Redirect("get_SpatialNeighbors")]
        NetOffice.VisioApi.IVSelection SpatialNeighbors(Int16 relation, Double tolerance, Int16 flags);

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <param name="x">Double x</param>
        /// <param name="y">Double y</param>
        /// <param name="relation">Int16 relation</param>
        /// <param name="tolerance">Double tolerance</param>
        /// <param name="flags">Int16 flags</param>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        NetOffice.VisioApi.IVSelection get_SpatialSearch(Double x, Double y, Int16 relation, Double tolerance, Int16 flags);

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// Alias for get_SpatialSearch
        /// </summary>
        /// <param name="x">Double x</param>
        /// <param name="y">Double y</param>
        /// <param name="relation">Int16 relation</param>
        /// <param name="tolerance">Double tolerance</param>
        /// <param name="flags">Int16 flags</param>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16), Redirect("get_SpatialSearch")]
        NetOffice.VisioApi.IVSelection SpatialSearch(Double x, Double y, Int16 relation, Double tolerance, Int16 flags);

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <param name="localeIndependentCellName">string localeIndependentCellName</param>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        NetOffice.VisioApi.IVCell get_CellsU(string localeIndependentCellName);

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// Alias for get_CellsU
        /// </summary>
        /// <param name="localeIndependentCellName">string localeIndependentCellName</param>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16), Redirect("get_CellsU")]
        NetOffice.VisioApi.IVCell CellsU(string localeIndependentCellName);

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        string NameU { get; set; }

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <param name="localeIndependentCellName">string localeIndependentCellName</param>
        /// <param name="fExistsLocally">Int16 fExistsLocally</param>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        Int16 get_CellExistsU(string localeIndependentCellName, Int16 fExistsLocally);

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// Alias for get_CellExistsU
        /// </summary>
        /// <param name="localeIndependentCellName">string localeIndependentCellName</param>
        /// <param name="fExistsLocally">Int16 fExistsLocally</param>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16), Redirect("get_CellExistsU")]
        Int16 CellExistsU(string localeIndependentCellName, Int16 fExistsLocally);

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <param name="localeSpecificCellName">string localeSpecificCellName</param>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        Int16 get_CellsRowIndex(string localeSpecificCellName);

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// Alias for get_CellsRowIndex
        /// </summary>
        /// <param name="localeSpecificCellName">string localeSpecificCellName</param>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16), Redirect("get_CellsRowIndex")]
        Int16 CellsRowIndex(string localeSpecificCellName);

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <param name="localeIndependentCellName">string localeIndependentCellName</param>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        Int16 get_CellsRowIndexU(string localeIndependentCellName);

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// Alias for get_CellsRowIndexU
        /// </summary>
        /// <param name="localeIndependentCellName">string localeIndependentCellName</param>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16), Redirect("get_CellsRowIndexU")]
        Int16 CellsRowIndexU(string localeIndependentCellName);

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        bool IsOpenForTextEdit { get; }

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        [BaseResult]
        NetOffice.VisioApi.IVShape RootShape { get; }

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        [BaseResult]
        NetOffice.VisioApi.IVShape MasterShape { get; }

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16), NativeResult]
        stdole.Picture Picture { get; }

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        byte[] ForeignData { get; }

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        Int32 Language { get; set; }

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        new Double AreaIU { get; }

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        new Double LengthIU { get; }

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        Int32 ContainingPageID { get; }

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        Int32 ContainingMasterID { get; }

        /// <summary>
        /// SupportByVersion Visio 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Visio", 12, 14, 15, 16)]
        [BaseResult]
        NetOffice.VisioApi.IVMaster DataGraphic { get; set; }

        /// <summary>
        /// SupportByVersion Visio 12, 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Visio", 12, 14, 15, 16)]
        bool IsDataGraphicCallout { get; }

        /// <summary>
        /// SupportByVersion Visio 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Visio", 14, 15, 16)]
        [BaseResult]
        NetOffice.VisioApi.IVContainerProperties ContainerProperties { get; }

        /// <summary>
        /// SupportByVersion Visio 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Visio", 14, 15, 16)]
        Int32[] MemberOfContainers { get; }

        /// <summary>
        /// SupportByVersion Visio 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Visio", 14, 15, 16)]
        bool IsCallout { get; }

        /// <summary>
        /// SupportByVersion Visio 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Visio", 14, 15, 16)]
        [BaseResult]
        NetOffice.VisioApi.IVShape CalloutTarget { get; set; }

        /// <summary>
        /// SupportByVersion Visio 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Visio", 14, 15, 16)]
        Int32[] CalloutsAssociated { get; }

        /// <summary>
        /// SupportByVersion Visio 15,16
        /// Get
        /// </summary>
        [SupportByVersion("Visio", 15, 16)]
        [BaseResult]
        NetOffice.VisioApi.IVComments Comments { get; }

        #endregion

        #region Methods

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// </summary>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        void VoidGroup();

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// </summary>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        void BringForward();

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// </summary>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        void BringToFront();

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// </summary>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        void ConvertToGroup();

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// </summary>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        void FlipHorizontal();

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// </summary>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        void FlipVertical();

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// </summary>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        void ReverseEnds();

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// </summary>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        void SendBackward();

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// </summary>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        void SendToBack();

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// </summary>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        void Rotate90();

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// </summary>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        void Ungroup();

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// </summary>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        void old_Copy();

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// </summary>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        void old_Cut();

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// </summary>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        void Delete();

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// </summary>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        void VoidDuplicate();

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="objectToDrop">object objectToDrop</param>
        /// <param name="xPos">Double xPos</param>
        /// <param name="yPos">Double yPos</param>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        [BaseResult]
        NetOffice.VisioApi.IVShape Drop(object objectToDrop, Double xPos, Double yPos);

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="section">Int16 section</param>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        Int16 AddSection(Int16 section);

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="section">Int16 section</param>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        void DeleteSection(Int16 section);

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="section">Int16 section</param>
        /// <param name="row">Int16 row</param>
        /// <param name="rowTag">Int16 rowTag</param>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        Int16 AddRow(Int16 section, Int16 row, Int16 rowTag);

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="section">Int16 section</param>
        /// <param name="row">Int16 row</param>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        void DeleteRow(Int16 section, Int16 row);

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="xPos">Double xPos</param>
        /// <param name="yPos">Double yPos</param>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        void SetCenter(Double xPos, Double yPos);

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="xPos">Double xPos</param>
        /// <param name="yPos">Double yPos</param>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        void SetBegin(Double xPos, Double yPos);

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="xPos">Double xPos</param>
        /// <param name="yPos">Double yPos</param>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        void SetEnd(Double xPos, Double yPos);

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="fileName">string fileName</param>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        void Export(string fileName);

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="section">Int16 section</param>
        /// <param name="rowName">string rowName</param>
        /// <param name="rowTag">Int16 rowTag</param>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        Int16 AddNamedRow(Int16 section, string rowName, Int16 rowTag);

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="section">Int16 section</param>
        /// <param name="row">Int16 row</param>
        /// <param name="rowTag">Int16 rowTag</param>
        /// <param name="rowCount">Int16 rowCount</param>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        Int16 AddRows(Int16 section, Int16 row, Int16 rowTag, Int16 rowCount);

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="xBegin">Double xBegin</param>
        /// <param name="yBegin">Double yBegin</param>
        /// <param name="xEnd">Double xEnd</param>
        /// <param name="yEnd">Double yEnd</param>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        [BaseResult]
        NetOffice.VisioApi.IVShape DrawLine(Double xBegin, Double yBegin, Double xEnd, Double yEnd);

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="x1">Double x1</param>
        /// <param name="y1">Double y1</param>
        /// <param name="x2">Double x2</param>
        /// <param name="y2">Double y2</param>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        [BaseResult]
        NetOffice.VisioApi.IVShape DrawRectangle(Double x1, Double y1, Double x2, Double y2);

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="x1">Double x1</param>
        /// <param name="y1">Double y1</param>
        /// <param name="x2">Double x2</param>
        /// <param name="y2">Double y2</param>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        [BaseResult]
        NetOffice.VisioApi.IVShape DrawOval(Double x1, Double y1, Double x2, Double y2);

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="xyArray">Double[] xyArray</param>
        /// <param name="tolerance">Double tolerance</param>
        /// <param name="flags">Int16 flags</param>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        NetOffice.VisioApi.IVShape DrawSpline(Double[] xyArray, Double tolerance, Int16 flags);

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="xyArray">Double[] xyArray</param>
        /// <param name="degree">Int16 degree</param>
        /// <param name="flags">Int16 flags</param>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        NetOffice.VisioApi.IVShape DrawBezier(Double[] xyArray, Int16 degree, Int16 flags);

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="xyArray">Double[] xyArray</param>
        /// <param name="flags">Int16 flags</param>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        NetOffice.VisioApi.IVShape DrawPolyline(Double[] xyArray, Int16 flags);

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="tolerance">Double tolerance</param>
        /// <param name="flags">Int16 flags</param>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        void FitCurve(Double tolerance, Int16 flags);

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="fileName">string fileName</param>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        [BaseResult]
        NetOffice.VisioApi.IVShape Import(string fileName);

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// </summary>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        void CenterDrawing();

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="fileName">string fileName</param>
        /// <param name="flags">Int16 flags</param>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        [BaseResult]
        NetOffice.VisioApi.IVShape InsertFromFile(string fileName, Int16 flags);

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="classOrProgID">string classOrProgID</param>
        /// <param name="flags">Int16 flags</param>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        [BaseResult]
        NetOffice.VisioApi.IVShape InsertObject(string classOrProgID, Int16 flags);

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// </summary>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        [BaseResult]
        NetOffice.VisioApi.IVWindow OpenDrawWindow();

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// </summary>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        [BaseResult]
        NetOffice.VisioApi.IVWindow OpenSheetWindow();

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="objectsToInstance">object[] objectsToInstance</param>
        /// <param name="xyArray">Double[] xyArray</param>
        /// <param name="iDArray">Int16[] iDArray</param>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        Int16 DropMany(object[] objectsToInstance, Double[] xyArray, out Int16[] iDArray);

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="sRCStream">Int16[] sRCStream</param>
        /// <param name="formulaArray">object[] formulaArray</param>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        void GetFormulas(Int16[] sRCStream, out object[] formulaArray);

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="sRCStream">Int16[] sRCStream</param>
        /// <param name="flags">Int16 flags</param>
        /// <param name="unitsNamesOrCodes">object[] unitsNamesOrCodes</param>
        /// <param name="resultArray">object[] resultArray</param>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        void GetResults(Int16[] sRCStream, Int16 flags, object[] unitsNamesOrCodes, out object[] resultArray);

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="sRCStream">Int16[] sRCStream</param>
        /// <param name="formulaArray">object[] formulaArray</param>
        /// <param name="flags">Int16 flags</param>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        Int16 SetFormulas(Int16[] sRCStream, object[] formulaArray, Int16 flags);

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="sRCStream">Int16[] sRCStream</param>
        /// <param name="unitsNamesOrCodes">object[] unitsNamesOrCodes</param>
        /// <param name="resultArray">object[] resultArray</param>
        /// <param name="flags">Int16 flags</param>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        Int16 SetResults(Int16[] sRCStream, object[] unitsNamesOrCodes, object[] resultArray, Int16 flags);

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// </summary>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        void Layout();

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="flags">Int16 flags</param>
        /// <param name="lpr8Left">Double lpr8Left</param>
        /// <param name="lpr8Bottom">Double lpr8Bottom</param>
        /// <param name="lpr8Right">Double lpr8Right</param>
        /// <param name="lpr8Top">Double lpr8Top</param>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        void BoundingBox(Int16 flags, out Double lpr8Left, out Double lpr8Bottom, out Double lpr8Right, out Double lpr8Top);

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="xPos">Double xPos</param>
        /// <param name="yPos">Double yPos</param>
        /// <param name="tolerance">Double tolerance</param>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        Int16 HitTest(Double xPos, Double yPos, Double tolerance);

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// </summary>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        [BaseResult]
        NetOffice.VisioApi.IVHyperlink AddHyperlink();

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="otherShape">NetOffice.VisioApi.IVShape otherShape</param>
        /// <param name="x">Double x</param>
        /// <param name="y">Double y</param>
        /// <param name="xprime">Double xprime</param>
        /// <param name="yprime">Double yprime</param>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        void TransformXYTo(NetOffice.VisioApi.IVShape otherShape, Double x, Double y, out Double xprime, out Double yprime);

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="otherShape">NetOffice.VisioApi.IVShape otherShape</param>
        /// <param name="x">Double x</param>
        /// <param name="y">Double y</param>
        /// <param name="xprime">Double xprime</param>
        /// <param name="yprime">Double yprime</param>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        void TransformXYFrom(NetOffice.VisioApi.IVShape otherShape, Double x, Double y, out Double xprime, out Double yprime);

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="x">Double x</param>
        /// <param name="y">Double y</param>
        /// <param name="xprime">Double xprime</param>
        /// <param name="yprime">Double yprime</param>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        void XYToPage(Double x, Double y, out Double xprime, out Double yprime);

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="x">Double x</param>
        /// <param name="y">Double y</param>
        /// <param name="xprime">Double xprime</param>
        /// <param name="yprime">Double yprime</param>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        void XYFromPage(Double x, Double y, out Double xprime, out Double yprime);

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// </summary>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        void UpdateAlignmentBox();

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="objectsToInstance">object[] objectsToInstance</param>
        /// <param name="xyArray">Double[] xyArray</param>
        /// <param name="iDArray">Int16[] iDArray</param>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        Int16 DropManyU(object[] objectsToInstance, Double[] xyArray, out Int16[] iDArray);

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="sRCStream">Int16[] sRCStream</param>
        /// <param name="formulaArray">object[] formulaArray</param>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        void GetFormulasU(Int16[] sRCStream, out object[] formulaArray);

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="degree">Int16 degree</param>
        /// <param name="flags">Int16 flags</param>
        /// <param name="xyArray">Double[] xyArray</param>
        /// <param name="knots">Double[] knots</param>
        /// <param name="weights">optional object weights</param>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        NetOffice.VisioApi.IVShape DrawNURBS(Int16 degree, Int16 flags, Double[] xyArray, Double[] knots, object weights);

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="degree">Int16 degree</param>
        /// <param name="flags">Int16 flags</param>
        /// <param name="xyArray">Double[] xyArray</param>
        /// <param name="knots">Double[] knots</param>
        [CustomMethod]
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        NetOffice.VisioApi.IVShape DrawNURBS(Int16 degree, Int16 flags, Double[] xyArray, Double[] knots);

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// </summary>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        [BaseResult]
        NetOffice.VisioApi.IVShape Group();

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// </summary>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        [BaseResult]
        NetOffice.VisioApi.IVShape Duplicate();

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// </summary>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        void SwapEnds();

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="flags">optional object flags</param>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        void Copy(object flags);

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// </summary>
        [CustomMethod]
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        void Copy();

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="flags">optional object flags</param>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        void Cut(object flags);

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// </summary>
        [CustomMethod]
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        void Cut();

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="flags">optional object flags</param>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        void Paste(object flags);

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// </summary>
        [CustomMethod]
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        void Paste();

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="format">Int32 format</param>
        /// <param name="link">optional object link</param>
        /// <param name="displayAsIcon">optional object displayAsIcon</param>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        void PasteSpecial(Int32 format, object link, object displayAsIcon);

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="format">Int32 format</param>
        [CustomMethod]
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        void PasteSpecial(Int32 format);

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="format">Int32 format</param>
        /// <param name="link">optional object link</param>
        [CustomMethod]
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        void PasteSpecial(Int32 format, object link);

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="selType">NetOffice.VisioApi.Enums.VisSelectionTypes selType</param>
        /// <param name="iterationMode">optional NetOffice.VisioApi.Enums.VisSelectMode IterationMode = 256</param>
        /// <param name="data">optional object data</param>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        [BaseResult]
        NetOffice.VisioApi.IVSelection CreateSelection(NetOffice.VisioApi.Enums.VisSelectionTypes selType, object iterationMode, object data);

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="selType">NetOffice.VisioApi.Enums.VisSelectionTypes selType</param>
        [CustomMethod]
        [BaseResult]
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        NetOffice.VisioApi.IVSelection CreateSelection(NetOffice.VisioApi.Enums.VisSelectionTypes selType);

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="selType">NetOffice.VisioApi.Enums.VisSelectionTypes selType</param>
        /// <param name="iterationMode">optional NetOffice.VisioApi.Enums.VisSelectMode IterationMode = 256</param>
        [CustomMethod]
        [BaseResult]
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        NetOffice.VisioApi.IVSelection CreateSelection(NetOffice.VisioApi.Enums.VisSelectionTypes selType, object iterationMode);

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="distance">Double distance</param>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        void Offset(Double distance);

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="type">Int16 type</param>
        /// <param name="xPos">Double xPos</param>
        /// <param name="yPos">Double yPos</param>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        [BaseResult]
        NetOffice.VisioApi.IVShape AddGuide(Int16 type, Double xPos, Double yPos);

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="xBegin">Double xBegin</param>
        /// <param name="yBegin">Double yBegin</param>
        /// <param name="xEnd">Double xEnd</param>
        /// <param name="yEnd">Double yEnd</param>
        /// <param name="xControl">Double xControl</param>
        /// <param name="yControl">Double yControl</param>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        [BaseResult]
        NetOffice.VisioApi.IVShape DrawArcByThreePoints(Double xBegin, Double yBegin, Double xEnd, Double yEnd, Double xControl, Double yControl);

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="xBegin">Double xBegin</param>
        /// <param name="yBegin">Double yBegin</param>
        /// <param name="xEnd">Double xEnd</param>
        /// <param name="yEnd">Double yEnd</param>
        /// <param name="sweepFlag">NetOffice.VisioApi.Enums.VisArcSweepFlags sweepFlag</param>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        [BaseResult]
        NetOffice.VisioApi.IVShape DrawQuarterArc(Double xBegin, Double yBegin, Double xEnd, Double yEnd, NetOffice.VisioApi.Enums.VisArcSweepFlags sweepFlag);

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="xCenter">Double xCenter</param>
        /// <param name="yCenter">Double yCenter</param>
        /// <param name="radius">Double radius</param>
        /// <param name="startAngle">optional Double StartAngle = 0</param>
        /// <param name="endAngle">optional Double EndAngle = 3.1415927410125732</param>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        [BaseResult]
        NetOffice.VisioApi.IVShape DrawCircularArc(Double xCenter, Double yCenter, Double radius, object startAngle, object endAngle);

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="xCenter">Double xCenter</param>
        /// <param name="yCenter">Double yCenter</param>
        /// <param name="radius">Double radius</param>
        [CustomMethod]
        [BaseResult]
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        NetOffice.VisioApi.IVShape DrawCircularArc(Double xCenter, Double yCenter, Double radius);

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="xCenter">Double xCenter</param>
        /// <param name="yCenter">Double yCenter</param>
        /// <param name="radius">Double radius</param>
        /// <param name="startAngle">optional Double StartAngle = 0</param>
        [CustomMethod]
        [BaseResult]
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        NetOffice.VisioApi.IVShape DrawCircularArc(Double xCenter, Double yCenter, Double radius, object startAngle);

        /// <summary>
        /// SupportByVersion Visio 12, 14, 15, 16
        /// </summary>
        /// <param name="dataRecordsetID">Int32 dataRecordsetID</param>
        /// <param name="rowID">Int32 rowID</param>
        /// <param name="applyDataGraphicAfterLink">optional bool ApplyDataGraphicAfterLink = true</param>
        [SupportByVersion("Visio", 12, 14, 15, 16)]
        void LinkToData(Int32 dataRecordsetID, Int32 rowID, object applyDataGraphicAfterLink);

        /// <summary>
        /// SupportByVersion Visio 12, 14, 15, 16
        /// </summary>
        /// <param name="dataRecordsetID">Int32 dataRecordsetID</param>
        /// <param name="rowID">Int32 rowID</param>
        [CustomMethod]
        [SupportByVersion("Visio", 12, 14, 15, 16)]
        void LinkToData(Int32 dataRecordsetID, Int32 rowID);

        /// <summary>
        /// SupportByVersion Visio 12, 14, 15, 16
        /// </summary>
        /// <param name="dataRecordsetID">Int32 dataRecordsetID</param>
        [SupportByVersion("Visio", 12, 14, 15, 16)]
        void BreakLinkToData(Int32 dataRecordsetID);

        /// <summary>
        /// SupportByVersion Visio 12, 14, 15, 16
        /// </summary>
        /// <param name="dataRecordsetID">Int32 dataRecordsetID</param>
        [SupportByVersion("Visio", 12, 14, 15, 16)]
        Int32 GetLinkedDataRow(Int32 dataRecordsetID);

        /// <summary>
        /// SupportByVersion Visio 12, 14, 15, 16
        /// </summary>
        /// <param name="dataRecordsetIDs">Int32[] dataRecordsetIDs</param>
        [SupportByVersion("Visio", 12, 14, 15, 16)]
        void GetLinkedDataRecordsetIDs(out Int32[] dataRecordsetIDs);

        /// <summary>
        /// SupportByVersion Visio 12, 14, 15, 16
        /// </summary>
        /// <param name="dataRecordsetID">Int32 dataRecordsetID</param>
        /// <param name="customPropertyIndices">Int32[] customPropertyIndices</param>
        [SupportByVersion("Visio", 12, 14, 15, 16)]
        void GetCustomPropertiesLinkedToData(Int32 dataRecordsetID, out Int32[] customPropertyIndices);

        /// <summary>
        /// SupportByVersion Visio 12, 14, 15, 16
        /// </summary>
        /// <param name="dataRecordsetID">Int32 dataRecordsetID</param>
        /// <param name="customPropertyIndex">Int32 customPropertyIndex</param>
        [SupportByVersion("Visio", 12, 14, 15, 16)]
        bool IsCustomPropertyLinked(Int32 dataRecordsetID, Int32 customPropertyIndex);

        /// <summary>
        /// SupportByVersion Visio 12, 14, 15, 16
        /// </summary>
        /// <param name="dataRecordsetID">Int32 dataRecordsetID</param>
        /// <param name="customPropertyIndex">Int32 customPropertyIndex</param>
        [SupportByVersion("Visio", 12, 14, 15, 16)]
        string GetCustomPropertyLinkedColumn(Int32 dataRecordsetID, Int32 customPropertyIndex);

        /// <summary>
        /// SupportByVersion Visio 12, 14, 15, 16
        /// </summary>
        /// <param name="toShape">NetOffice.VisioApi.IVShape toShape</param>
        /// <param name="placementDir">NetOffice.VisioApi.Enums.VisAutoConnectDir placementDir</param>
        /// <param name="connector">optional object Connector = null (Nothing in visual basic)</param>
        [SupportByVersion("Visio", 12, 14, 15, 16)]
        void AutoConnect(NetOffice.VisioApi.IVShape toShape, NetOffice.VisioApi.Enums.VisAutoConnectDir placementDir, object connector);

        /// <summary>
        /// SupportByVersion Visio 12, 14, 15, 16
        /// </summary>
        /// <param name="toShape">NetOffice.VisioApi.IVShape toShape</param>
        /// <param name="placementDir">NetOffice.VisioApi.Enums.VisAutoConnectDir placementDir</param>
        [CustomMethod]
        [SupportByVersion("Visio", 12, 14, 15, 16)]
        void AutoConnect(NetOffice.VisioApi.IVShape toShape, NetOffice.VisioApi.Enums.VisAutoConnectDir placementDir);

        /// <summary>
        /// SupportByVersion Visio 14, 15, 16
        /// </summary>
        /// <param name="category">string category</param>
        [SupportByVersion("Visio", 14, 15, 16)]
        bool HasCategory(string category);

        /// <summary>
        /// SupportByVersion Visio 14, 15, 16
        /// </summary>
        /// <param name="flags">NetOffice.VisioApi.Enums.VisConnectedShapesFlags flags</param>
        /// <param name="categoryFilter">string categoryFilter</param>
        [SupportByVersion("Visio", 14, 15, 16)]
        Int32[] ConnectedShapes(NetOffice.VisioApi.Enums.VisConnectedShapesFlags flags, string categoryFilter);

        /// <summary>
        /// SupportByVersion Visio 14, 15, 16
        /// </summary>
        /// <param name="flags">NetOffice.VisioApi.Enums.VisGluedShapesFlags flags</param>
        /// <param name="categoryFilter">string categoryFilter</param>
        /// <param name="pOtherConnectedShape">optional NetOffice.VisioApi.IVShape pOtherConnectedShape</param>
        [SupportByVersion("Visio", 14, 15, 16)]
        Int32[] GluedShapes(NetOffice.VisioApi.Enums.VisGluedShapesFlags flags, string categoryFilter, object pOtherConnectedShape);

        /// <summary>
        /// SupportByVersion Visio 14, 15, 16
        /// </summary>
        /// <param name="flags">NetOffice.VisioApi.Enums.VisGluedShapesFlags flags</param>
        /// <param name="categoryFilter">string categoryFilter</param>
        [CustomMethod]
        [SupportByVersion("Visio", 14, 15, 16)]
        Int32[] GluedShapes(NetOffice.VisioApi.Enums.VisGluedShapesFlags flags, string categoryFilter);

        /// <summary>
        /// SupportByVersion Visio 14, 15, 16
        /// </summary>
        /// <param name="connectorEnd">NetOffice.VisioApi.Enums.VisConnectorEnds connectorEnd</param>
        /// <param name="offsetX">Double offsetX</param>
        /// <param name="offsetY">Double offsetY</param>
        /// <param name="units">NetOffice.VisioApi.Enums.VisUnitCodes units</param>
        [SupportByVersion("Visio", 14, 15, 16)]
        void Disconnect(NetOffice.VisioApi.Enums.VisConnectorEnds connectorEnd, Double offsetX, Double offsetY, NetOffice.VisioApi.Enums.VisUnitCodes units);

        /// <summary>
        /// SupportByVersion Visio 14, 15, 16
        /// </summary>
        /// <param name="direction">NetOffice.VisioApi.Enums.VisResizeDirection direction</param>
        /// <param name="distance">Double distance</param>
        /// <param name="unitCode">NetOffice.VisioApi.Enums.VisUnitCodes unitCode</param>
        [SupportByVersion("Visio", 14, 15, 16)]
        void Resize(NetOffice.VisioApi.Enums.VisResizeDirection direction, Double distance, NetOffice.VisioApi.Enums.VisUnitCodes unitCode);

        /// <summary>
        /// SupportByVersion Visio 14, 15, 16
        /// </summary>
        [SupportByVersion("Visio", 14, 15, 16)]
        void AddToContainers();

        /// <summary>
        /// SupportByVersion Visio 14, 15, 16
        /// </summary>
        [SupportByVersion("Visio", 14, 15, 16)]
        void RemoveFromContainers();

        /// <summary>
        /// SupportByVersion Visio 14, 15, 16
        /// </summary>
        [SupportByVersion("Visio", 14, 15, 16)]
        [BaseResult]
        NetOffice.VisioApi.IVPage CreateSubProcess();

        /// <summary>
        /// SupportByVersion Visio 14, 15, 16
        /// </summary>
        /// <param name="page">NetOffice.VisioApi.IVPage page</param>
        /// <param name="objectToDrop">object objectToDrop</param>
        /// <param name="newShape">optional NetOffice.VisioApi.IVShape NewShape = 0</param>
        [SupportByVersion("Visio", 14, 15, 16)]
        [BaseResult]
        NetOffice.VisioApi.IVSelection MoveToSubprocess(NetOffice.VisioApi.IVPage page, object objectToDrop, object newShape);

        /// <summary>
        /// SupportByVersion Visio 14, 15, 16
        /// </summary>
        /// <param name="page">NetOffice.VisioApi.IVPage page</param>
        /// <param name="objectToDrop">object objectToDrop</param>
        [CustomMethod]
        [BaseResult]
        [SupportByVersion("Visio", 14, 15, 16)]
        NetOffice.VisioApi.IVSelection MoveToSubprocess(NetOffice.VisioApi.IVPage page, object objectToDrop);

        /// <summary>
        /// SupportByVersion Visio 14, 15, 16
        /// </summary>
        /// <param name="delFlags">Int32 delFlags</param>
        [SupportByVersion("Visio", 14, 15, 16)]
        void DeleteEx(Int32 delFlags);

        /// <summary>
        /// SupportByVersion Visio 15,16
        /// </summary>
        /// <param name="masterOrMasterShortcutToDrop">object masterOrMasterShortcutToDrop</param>
        /// <param name="replaceFlags">optional Int32 ReplaceFlags = 0</param>
        [SupportByVersion("Visio", 15, 16)]
        [BaseResult]
        NetOffice.VisioApi.IVShape ReplaceShape(object masterOrMasterShortcutToDrop, object replaceFlags);

        /// <summary>
        /// SupportByVersion Visio 15,16
        /// </summary>
        /// <param name="masterOrMasterShortcutToDrop">object masterOrMasterShortcutToDrop</param>
        [CustomMethod]
        [BaseResult]
        [SupportByVersion("Visio", 15, 16)]
        NetOffice.VisioApi.IVShape ReplaceShape(object masterOrMasterShortcutToDrop);

        /// <summary>
        /// SupportByVersion Visio 15,16
        /// </summary>
        /// <param name="lineMatrix">NetOffice.VisioApi.Enums.VisQuickStyleMatrixIndices lineMatrix</param>
        /// <param name="fillMatrix">NetOffice.VisioApi.Enums.VisQuickStyleMatrixIndices fillMatrix</param>
        /// <param name="effectsMatrix">NetOffice.VisioApi.Enums.VisQuickStyleMatrixIndices effectsMatrix</param>
        /// <param name="fontMatrix">NetOffice.VisioApi.Enums.VisQuickStyleMatrixIndices fontMatrix</param>
        /// <param name="lineColor">NetOffice.VisioApi.Enums.VisQuickStyleColors lineColor</param>
        /// <param name="fillColor">NetOffice.VisioApi.Enums.VisQuickStyleColors fillColor</param>
        /// <param name="shadowColor">NetOffice.VisioApi.Enums.VisQuickStyleColors shadowColor</param>
        /// <param name="fontColor">NetOffice.VisioApi.Enums.VisQuickStyleColors fontColor</param>
        [SupportByVersion("Visio", 15, 16)]
        void SetQuickStyle(NetOffice.VisioApi.Enums.VisQuickStyleMatrixIndices lineMatrix, NetOffice.VisioApi.Enums.VisQuickStyleMatrixIndices fillMatrix, NetOffice.VisioApi.Enums.VisQuickStyleMatrixIndices effectsMatrix, NetOffice.VisioApi.Enums.VisQuickStyleMatrixIndices fontMatrix, NetOffice.VisioApi.Enums.VisQuickStyleColors lineColor, NetOffice.VisioApi.Enums.VisQuickStyleColors fillColor, NetOffice.VisioApi.Enums.VisQuickStyleColors shadowColor, NetOffice.VisioApi.Enums.VisQuickStyleColors fontColor);

        /// <summary>
        /// SupportByVersion Visio 15,16
        /// </summary>
        /// <param name="fileName">string fileName</param>
        /// <param name="changePictureFlags">optional Int32 ChangePictureFlags = 0</param>
        [SupportByVersion("Visio", 15, 16)]
        Double ChangePicture(string fileName, object changePictureFlags);

        /// <summary>
        /// SupportByVersion Visio 15,16
        /// </summary>
        /// <param name="fileName">string fileName</param>
        [CustomMethod]
        [SupportByVersion("Visio", 15, 16)]
        Double ChangePicture(string fileName);

        #endregion
    }
}
