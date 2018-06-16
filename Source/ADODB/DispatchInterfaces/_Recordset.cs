using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.ADODBApi
{
	/// <summary>
	/// DispatchInterface _Recordset 
	/// SupportByVersion ADODB, 2.1,2.5
	/// </summary>
	[SupportByVersion("ADODB", 2.1,2.5)]
	[EntityType(EntityType.IsDispatchInterface), BaseType]
	[TypeId("00000555-0000-0010-8000-00AA006D2EA4")]
    [CoClassSource(typeof(NetOffice.ADODBApi.Recordset))]
    public interface _Recordset : Recordset21
	{
		#region Properties

		/// <summary>
		/// SupportByVersion ADODB 2.1, 2.5
		/// Get
		/// </summary>
		[SupportByVersion("ADODB", 2.1,2.5)]
		new NetOffice.ADODBApi.Properties Properties { get; }

		/// <summary>
		/// SupportByVersion ADODB 2.1, 2.5
		/// Get/Set
		/// </summary>
		[SupportByVersion("ADODB", 2.1,2.5)]
        new NetOffice.ADODBApi.Enums.PositionEnum AbsolutePosition { get; set; }

		/// <summary>
		/// SupportByVersion ADODB 2.1, 2.5
		/// Get/Set
		/// Unknown COM Proxy
		/// </summary>
		[SupportByVersion("ADODB", 2.1,2.5), ProxyResult]
        new object ActiveConnection { get; set; }

		/// <summary>
		/// SupportByVersion ADODB 2.1, 2.5
		/// Get
		/// </summary>
		[SupportByVersion("ADODB", 2.1,2.5)]
        new bool BOF { get; }

		/// <summary>
		/// SupportByVersion ADODB 2.1, 2.5
		/// Get/Set
		/// </summary>
		[SupportByVersion("ADODB", 2.1,2.5)]
        new object Bookmark { get; set; }

		/// <summary>
		/// SupportByVersion ADODB 2.1, 2.5
		/// Get/Set
		/// </summary>
		[SupportByVersion("ADODB", 2.1,2.5)]
        new Int32 CacheSize { get; set; }

		/// <summary>
		/// SupportByVersion ADODB 2.1, 2.5
		/// Get/Set
		/// </summary>
		[SupportByVersion("ADODB", 2.1,2.5)]
        new NetOffice.ADODBApi.Enums.CursorTypeEnum CursorType { get; set; }

		/// <summary>
		/// SupportByVersion ADODB 2.1, 2.5
		/// Get
		/// </summary>
		[SupportByVersion("ADODB", 2.1,2.5)]
        new bool EOF { get; }

		/// <summary>
		/// SupportByVersion ADODB 2.1, 2.5
		/// Get
		/// </summary>
		[SupportByVersion("ADODB", 2.1,2.5)]
        new NetOffice.ADODBApi.Fields Fields { get; }

		/// <summary>
		/// SupportByVersion ADODB 2.1, 2.5
		/// Get/Set
		/// </summary>
		[SupportByVersion("ADODB", 2.1,2.5)]
        new NetOffice.ADODBApi.Enums.LockTypeEnum LockType { get; set; }

		/// <summary>
		/// SupportByVersion ADODB 2.1, 2.5
		/// Get/Set
		/// </summary>
		[SupportByVersion("ADODB", 2.1,2.5)]
        new Int32 MaxRecords { get; set; }

		/// <summary>
		/// SupportByVersion ADODB 2.1, 2.5
		/// Get
		/// </summary>
		[SupportByVersion("ADODB", 2.1,2.5)]
        new Int32 RecordCount { get; }

		/// <summary>
		/// SupportByVersion ADODB 2.1, 2.5
		/// Get/Set
		/// Unknown COM Proxy
		/// </summary>
		[SupportByVersion("ADODB", 2.1,2.5), ProxyResult]
        new object Source { get; set; }

		/// <summary>
		/// SupportByVersion ADODB 2.1, 2.5
		/// Get/Set
		/// </summary>
		[SupportByVersion("ADODB", 2.1,2.5)]
        new NetOffice.ADODBApi.Enums.PositionEnum AbsolutePage { get; set; }

		/// <summary>
		/// SupportByVersion ADODB 2.1, 2.5
		/// Get
		/// </summary>
		[SupportByVersion("ADODB", 2.1,2.5)]
        new NetOffice.ADODBApi.Enums.EditModeEnum EditMode { get; }

		/// <summary>
		/// SupportByVersion ADODB 2.1, 2.5
		/// Get/Set
		/// </summary>
		[SupportByVersion("ADODB", 2.1,2.5)]
        new object Filter { get; set; }

		/// <summary>
		/// SupportByVersion ADODB 2.1, 2.5
		/// Get
		/// </summary>
		[SupportByVersion("ADODB", 2.1,2.5)]
        new Int32 PageCount { get; }

		/// <summary>
		/// SupportByVersion ADODB 2.1, 2.5
		/// Get/Set
		/// </summary>
		[SupportByVersion("ADODB", 2.1,2.5)]
        new Int32 PageSize { get; set; }

		/// <summary>
		/// SupportByVersion ADODB 2.1, 2.5
		/// Get/Set
		/// </summary>
		[SupportByVersion("ADODB", 2.1,2.5)]
        new string Sort { get; set; }

		/// <summary>
		/// SupportByVersion ADODB 2.1, 2.5
		/// Get
		/// </summary>
		[SupportByVersion("ADODB", 2.1,2.5)]
        new Int32 Status { get; }

		/// <summary>
		/// SupportByVersion ADODB 2.1, 2.5
		/// Get
		/// </summary>
		[SupportByVersion("ADODB", 2.1,2.5)]
        new Int32 State { get; }

		/// <summary>
		/// SupportByVersion ADODB 2.1, 2.5
		/// Get/Set
		/// </summary>
		[SupportByVersion("ADODB", 2.1,2.5)]
        new NetOffice.ADODBApi.Enums.CursorLocationEnum CursorLocation { get; set; }

		/// <summary>
		/// SupportByVersion ADODB 2.1, 2.5
		/// Get/Set
		/// </summary>
		/// <param name="index">object index</param>
		[SupportByVersion("ADODB", 2.1,2.5)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        new object get_Collect(object index);

        /// <summary>
        /// SupportByVersion ADODB 2.1, 2.5
        /// Get/Set
        /// </summary>
        /// <param name="index">object index</param>
        /// <param name="value">object value</param>
        [SupportByVersion("ADODB", 2.1,2.5)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        new void set_Collect(object index, object value);

		/// <summary>
		/// SupportByVersion ADODB 2.1, 2.5
		/// Alias for get_Collect
		/// </summary>
		/// <param name="index">object index</param>
		[SupportByVersion("ADODB", 2.1,2.5), Redirect("get_Collect")]
        new object Collect(object index);

		/// <summary>
		/// SupportByVersion ADODB 2.1, 2.5
		/// Get/Set
		/// </summary>
		[SupportByVersion("ADODB", 2.1,2.5)]
        new NetOffice.ADODBApi.Enums.MarshalOptionsEnum MarshalOptions { get; set; }

		/// <summary>
		/// SupportByVersion ADODB 2.1
		/// Get/Set
		/// </summary>
		[SupportByVersion("ADODB", 2.1)]
        new string Index { get; set; }

		/// <summary>
		/// SupportByVersion ADODB 2.5
		/// Get/Set
		/// Unknown COM Proxy
		/// </summary>
		[SupportByVersion("ADODB", 2.5), ProxyResult]
        new object DataSource { get; set; }

		/// <summary>
		/// SupportByVersion ADODB 2.5
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		[SupportByVersion("ADODB", 2.5), ProxyResult]
        new object ActiveCommand { get; }

		/// <summary>
		/// SupportByVersion ADODB 2.5
		/// Get/Set
		/// </summary>
		[SupportByVersion("ADODB", 2.5)]
        new bool StayInSync { get; set; }

		/// <summary>
		/// SupportByVersion ADODB 2.5
		/// Get/Set
		/// </summary>
		[SupportByVersion("ADODB", 2.5)]
        new string DataMember { get; set; }

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion ADODB 2.1, 2.5
		/// </summary>
		/// <param name="fieldList">optional object fieldList</param>
		/// <param name="values">optional object values</param>
		[SupportByVersion("ADODB", 2.1,2.5)]
        new void AddNew(object fieldList, object values);

		/// <summary>
		/// SupportByVersion ADODB 2.1, 2.5
		/// </summary>
		[CustomMethod]
		[SupportByVersion("ADODB", 2.1,2.5)]
        new void AddNew();

		/// <summary>
		/// SupportByVersion ADODB 2.1, 2.5
		/// </summary>
		/// <param name="fieldList">optional object fieldList</param>
		[CustomMethod]
		[SupportByVersion("ADODB", 2.1,2.5)]
        new void AddNew(object fieldList);

		/// <summary>
		/// SupportByVersion ADODB 2.1, 2.5
		/// </summary>
		[SupportByVersion("ADODB", 2.1,2.5)]
        new void CancelUpdate();

		/// <summary>
		/// SupportByVersion ADODB 2.1, 2.5
		/// </summary>
		[SupportByVersion("ADODB", 2.1,2.5)]
        new void Close();

		/// <summary>
		/// SupportByVersion ADODB 2.1, 2.5
		/// </summary>
		/// <param name="affectRecords">optional NetOffice.ADODBApi.Enums.AffectEnum AffectRecords = 1</param>
		[SupportByVersion("ADODB", 2.1,2.5)]
        new void Delete(object affectRecords);

		/// <summary>
		/// SupportByVersion ADODB 2.1, 2.5
		/// </summary>
		[CustomMethod]
		[SupportByVersion("ADODB", 2.1,2.5)]
        new void Delete();

		/// <summary>
		/// SupportByVersion ADODB 2.1, 2.5
		/// </summary>
		/// <param name="rows">optional Int32 Rows = -1</param>
		/// <param name="start">optional object start</param>
		/// <param name="fields">optional object fields</param>
		[SupportByVersion("ADODB", 2.1,2.5)]
        new object GetRows(object rows, object start, object fields);

		/// <summary>
		/// SupportByVersion ADODB 2.1, 2.5
		/// </summary>
		[CustomMethod]
		[SupportByVersion("ADODB", 2.1,2.5)]
        new object GetRows();

		/// <summary>
		/// SupportByVersion ADODB 2.1, 2.5
		/// </summary>
		/// <param name="rows">optional Int32 Rows = -1</param>
		[CustomMethod]
		[SupportByVersion("ADODB", 2.1,2.5)]
        new object GetRows(object rows);

		/// <summary>
		/// SupportByVersion ADODB 2.1, 2.5
		/// </summary>
		/// <param name="rows">optional Int32 Rows = -1</param>
		/// <param name="start">optional object start</param>
		[CustomMethod]
		[SupportByVersion("ADODB", 2.1,2.5)]
        new object GetRows(object rows, object start);

		/// <summary>
		/// SupportByVersion ADODB 2.1, 2.5
		/// </summary>
		/// <param name="numRecords">Int32 numRecords</param>
		/// <param name="start">optional object start</param>
		[SupportByVersion("ADODB", 2.1,2.5)]
        new void Move(Int32 numRecords, object start);

		/// <summary>
		/// SupportByVersion ADODB 2.1, 2.5
		/// </summary>
		/// <param name="numRecords">Int32 numRecords</param>
		[CustomMethod]
		[SupportByVersion("ADODB", 2.1,2.5)]
        new void Move(Int32 numRecords);

		/// <summary>
		/// SupportByVersion ADODB 2.1, 2.5
		/// </summary>
		[SupportByVersion("ADODB", 2.1,2.5)]
        new void MoveNext();

		/// <summary>
		/// SupportByVersion ADODB 2.1, 2.5
		/// </summary>
		[SupportByVersion("ADODB", 2.1,2.5)]
        new void MovePrevious();

		/// <summary>
		/// SupportByVersion ADODB 2.1, 2.5
		/// </summary>
		[SupportByVersion("ADODB", 2.1,2.5)]
        new void MoveFirst();

		/// <summary>
		/// SupportByVersion ADODB 2.1, 2.5
		/// </summary>
		[SupportByVersion("ADODB", 2.1,2.5)]
        new void MoveLast();

		/// <summary>
		/// SupportByVersion ADODB 2.1, 2.5
		/// </summary>
		/// <param name="source">optional object source</param>
		/// <param name="activeConnection">optional object activeConnection</param>
		/// <param name="cursorType">optional NetOffice.ADODBApi.Enums.CursorTypeEnum CursorType = -1</param>
		/// <param name="lockType">optional NetOffice.ADODBApi.Enums.LockTypeEnum LockType = -1</param>
		/// <param name="options">optional Int32 Options = -1</param>
		[SupportByVersion("ADODB", 2.1,2.5)]
        new void Open(object source, object activeConnection, object cursorType, object lockType, object options);

		/// <summary>
		/// SupportByVersion ADODB 2.1, 2.5
		/// </summary>
		[CustomMethod]
		[SupportByVersion("ADODB", 2.1,2.5)]
        new void Open();

		/// <summary>
		/// SupportByVersion ADODB 2.1, 2.5
		/// </summary>
		/// <param name="source">optional object source</param>
		[CustomMethod]
		[SupportByVersion("ADODB", 2.1,2.5)]
        new void Open(object source);

		/// <summary>
		/// SupportByVersion ADODB 2.1, 2.5
		/// </summary>
		/// <param name="source">optional object source</param>
		/// <param name="activeConnection">optional object activeConnection</param>
		[CustomMethod]
		[SupportByVersion("ADODB", 2.1,2.5)]
        new void Open(object source, object activeConnection);

		/// <summary>
		/// SupportByVersion ADODB 2.1, 2.5
		/// </summary>
		/// <param name="source">optional object source</param>
		/// <param name="activeConnection">optional object activeConnection</param>
		/// <param name="cursorType">optional NetOffice.ADODBApi.Enums.CursorTypeEnum CursorType = -1</param>
		[CustomMethod]
		[SupportByVersion("ADODB", 2.1,2.5)]
        new void Open(object source, object activeConnection, object cursorType);

		/// <summary>
		/// SupportByVersion ADODB 2.1, 2.5
		/// </summary>
		/// <param name="source">optional object source</param>
		/// <param name="activeConnection">optional object activeConnection</param>
		/// <param name="cursorType">optional NetOffice.ADODBApi.Enums.CursorTypeEnum CursorType = -1</param>
		/// <param name="lockType">optional NetOffice.ADODBApi.Enums.LockTypeEnum LockType = -1</param>
		[CustomMethod]
		[SupportByVersion("ADODB", 2.1,2.5)]
        new void Open(object source, object activeConnection, object cursorType, object lockType);

		/// <summary>
		/// SupportByVersion ADODB 2.1, 2.5
		/// </summary>
		/// <param name="options">optional Int32 Options = -1</param>
		[SupportByVersion("ADODB", 2.1,2.5)]
        new void Requery(object options);

		/// <summary>
		/// SupportByVersion ADODB 2.1, 2.5
		/// </summary>
		[CustomMethod]
		[SupportByVersion("ADODB", 2.1,2.5)]
        new void Requery();

		/// <summary>
		/// SupportByVersion ADODB 2.1, 2.5
		/// </summary>
		/// <param name="affectRecords">optional NetOffice.ADODBApi.Enums.AffectEnum AffectRecords = 3</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("ADODB", 2.1,2.5)]
        new  void _xResync(object affectRecords);

		/// <summary>
		/// SupportByVersion ADODB 2.1, 2.5
		/// </summary>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("ADODB", 2.1,2.5)]
        new void _xResync();

		/// <summary>
		/// SupportByVersion ADODB 2.1, 2.5
		/// </summary>
		/// <param name="fields">optional object fields</param>
		/// <param name="values">optional object values</param>
		[SupportByVersion("ADODB", 2.1,2.5)]
        new void Update(object fields, object values);

		/// <summary>
		/// SupportByVersion ADODB 2.1, 2.5
		/// </summary>
		[CustomMethod]
		[SupportByVersion("ADODB", 2.1,2.5)]
        new void Update();

		/// <summary>
		/// SupportByVersion ADODB 2.1, 2.5
		/// </summary>
		/// <param name="fields">optional object fields</param>
		[CustomMethod]
		[SupportByVersion("ADODB", 2.1,2.5)]
        new void Update(object fields);

		/// <summary>
		/// SupportByVersion ADODB 2.1, 2.5
		/// </summary>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[BaseResult]
		[SupportByVersion("ADODB", 2.1,2.5)]
        new NetOffice.ADODBApi._Recordset _xClone();

		/// <summary>
		/// SupportByVersion ADODB 2.1, 2.5
		/// </summary>
		/// <param name="affectRecords">optional NetOffice.ADODBApi.Enums.AffectEnum AffectRecords = 3</param>
		[SupportByVersion("ADODB", 2.1,2.5)]
        new void UpdateBatch(object affectRecords);

		/// <summary>
		/// SupportByVersion ADODB 2.1, 2.5
		/// </summary>
		[CustomMethod]
		[SupportByVersion("ADODB", 2.1,2.5)]
        new void UpdateBatch();

		/// <summary>
		/// SupportByVersion ADODB 2.1, 2.5
		/// </summary>
		/// <param name="affectRecords">optional NetOffice.ADODBApi.Enums.AffectEnum AffectRecords = 3</param>
		[SupportByVersion("ADODB", 2.1,2.5)]
        new void CancelBatch(object affectRecords);

		/// <summary>
		/// SupportByVersion ADODB 2.1, 2.5
		/// </summary>
		[CustomMethod]
		[SupportByVersion("ADODB", 2.1,2.5)]
        new void CancelBatch();

		/// <summary>
		/// SupportByVersion ADODB 2.1, 2.5
		/// </summary>
		/// <param name="recordsAffected">optional object recordsAffected</param>
		[SupportByVersion("ADODB", 2.1,2.5)]
		[BaseResult]
        new NetOffice.ADODBApi._Recordset NextRecordset(object recordsAffected);

		/// <summary>
		/// SupportByVersion ADODB 2.1, 2.5
		/// </summary>
		[CustomMethod]
		[BaseResult]
		[SupportByVersion("ADODB", 2.1,2.5)]
        new NetOffice.ADODBApi._Recordset NextRecordset();

		/// <summary>
		/// SupportByVersion ADODB 2.1, 2.5
		/// </summary>
		/// <param name="cursorOptions">NetOffice.ADODBApi.Enums.CursorOptionEnum cursorOptions</param>
		[SupportByVersion("ADODB", 2.1,2.5)]
        new bool Supports(NetOffice.ADODBApi.Enums.CursorOptionEnum cursorOptions);

		/// <summary>
		/// SupportByVersion ADODB 2.1, 2.5
		/// </summary>
		/// <param name="criteria">string criteria</param>
		/// <param name="skipRecords">optional Int32 SkipRecords = 0</param>
		/// <param name="searchDirection">optional NetOffice.ADODBApi.Enums.SearchDirectionEnum SearchDirection = 1</param>
		/// <param name="start">optional object start</param>
		[SupportByVersion("ADODB", 2.1,2.5)]
        new void Find(string criteria, object skipRecords, object searchDirection, object start);

		/// <summary>
		/// SupportByVersion ADODB 2.1, 2.5
		/// </summary>
		/// <param name="criteria">string criteria</param>
		[CustomMethod]
		[SupportByVersion("ADODB", 2.1,2.5)]
        new void Find(string criteria);

		/// <summary>
		/// SupportByVersion ADODB 2.1, 2.5
		/// </summary>
		/// <param name="criteria">string criteria</param>
		/// <param name="skipRecords">optional Int32 SkipRecords = 0</param>
		[CustomMethod]
		[SupportByVersion("ADODB", 2.1,2.5)]
        new void Find(string criteria, object skipRecords);

		/// <summary>
		/// SupportByVersion ADODB 2.1, 2.5
		/// </summary>
		/// <param name="criteria">string criteria</param>
		/// <param name="skipRecords">optional Int32 SkipRecords = 0</param>
		/// <param name="searchDirection">optional NetOffice.ADODBApi.Enums.SearchDirectionEnum SearchDirection = 1</param>
		[CustomMethod]
		[SupportByVersion("ADODB", 2.1,2.5)]
        new void Find(string criteria, object skipRecords, object searchDirection);

		/// <summary>
		/// SupportByVersion ADODB 2.1
		/// </summary>
		/// <param name="keyValues">object keyValues</param>
		/// <param name="seekOption">optional NetOffice.ADODBApi.Enums.SeekEnum SeekOption = 1</param>
		[SupportByVersion("ADODB", 2.1)]
        new void Seek(object keyValues, object seekOption);

		/// <summary>
		/// SupportByVersion ADODB 2.1
		/// </summary>
		/// <param name="keyValues">object keyValues</param>
		[CustomMethod]
		[SupportByVersion("ADODB", 2.1)]
        new void Seek(object keyValues);

		/// <summary>
		/// SupportByVersion ADODB 2.5
		/// </summary>
		[SupportByVersion("ADODB", 2.5)]
        new void Cancel();

		/// <summary>
		/// SupportByVersion ADODB 2.5
		/// </summary>
		/// <param name="stringFormat">optional NetOffice.ADODBApi.Enums.StringFormatEnum StringFormat = 2</param>
		/// <param name="numRows">optional Int32 NumRows = -1</param>
		/// <param name="columnDelimeter">optional string columnDelimeter</param>
		/// <param name="rowDelimeter">optional string rowDelimeter</param>
		/// <param name="nullExpr">optional string nullExpr</param>
		[SupportByVersion("ADODB", 2.5)]
        new string GetString(object stringFormat, object numRows, object columnDelimeter, object rowDelimeter, object nullExpr);

		/// <summary>
		/// SupportByVersion ADODB 2.5
		/// </summary>
		[CustomMethod]
		[SupportByVersion("ADODB", 2.5)]
        new string GetString();

		/// <summary>
		/// SupportByVersion ADODB 2.5
		/// </summary>
		/// <param name="stringFormat">optional NetOffice.ADODBApi.Enums.StringFormatEnum StringFormat = 2</param>
		[CustomMethod]
		[SupportByVersion("ADODB", 2.5)]
        new string GetString(object stringFormat);

		/// <summary>
		/// SupportByVersion ADODB 2.5
		/// </summary>
		/// <param name="stringFormat">optional NetOffice.ADODBApi.Enums.StringFormatEnum StringFormat = 2</param>
		/// <param name="numRows">optional Int32 NumRows = -1</param>
		[CustomMethod]
		[SupportByVersion("ADODB", 2.5)]
        new string GetString(object stringFormat, object numRows);

		/// <summary>
		/// SupportByVersion ADODB 2.5
		/// </summary>
		/// <param name="stringFormat">optional NetOffice.ADODBApi.Enums.StringFormatEnum StringFormat = 2</param>
		/// <param name="numRows">optional Int32 NumRows = -1</param>
		/// <param name="columnDelimeter">optional string columnDelimeter</param>
		[CustomMethod]
		[SupportByVersion("ADODB", 2.5)]
        new string GetString(object stringFormat, object numRows, object columnDelimeter);

		/// <summary>
		/// SupportByVersion ADODB 2.5
		/// </summary>
		/// <param name="stringFormat">optional NetOffice.ADODBApi.Enums.StringFormatEnum StringFormat = 2</param>
		/// <param name="numRows">optional Int32 NumRows = -1</param>
		/// <param name="columnDelimeter">optional string columnDelimeter</param>
		/// <param name="rowDelimeter">optional string rowDelimeter</param>
		[CustomMethod]
		[SupportByVersion("ADODB", 2.5)]
        new string GetString(object stringFormat, object numRows, object columnDelimeter, object rowDelimeter);

		/// <summary>
		/// SupportByVersion ADODB 2.5
		/// </summary>
		/// <param name="bookmark1">object bookmark1</param>
		/// <param name="bookmark2">object bookmark2</param>
		[SupportByVersion("ADODB", 2.5)]
        new NetOffice.ADODBApi.Enums.CompareEnum CompareBookmarks(object bookmark1, object bookmark2);

		/// <summary>
		/// SupportByVersion ADODB 2.5
		/// </summary>
		/// <param name="lockType">optional NetOffice.ADODBApi.Enums.LockTypeEnum LockType = -1</param>
		[SupportByVersion("ADODB", 2.5)]
		[BaseResult]
        new NetOffice.ADODBApi._Recordset Clone(object lockType);

		/// <summary>
		/// SupportByVersion ADODB 2.5
		/// </summary>
		[CustomMethod]
		[BaseResult]
		[SupportByVersion("ADODB", 2.5)]
        new NetOffice.ADODBApi._Recordset Clone();

		/// <summary>
		/// SupportByVersion ADODB 2.5
		/// </summary>
		/// <param name="affectRecords">optional NetOffice.ADODBApi.Enums.AffectEnum AffectRecords = 3</param>
		/// <param name="resyncValues">optional NetOffice.ADODBApi.Enums.ResyncEnum ResyncValues = 2</param>
		[SupportByVersion("ADODB", 2.5)]
        new void Resync(object affectRecords, object resyncValues);

		/// <summary>
		/// SupportByVersion ADODB 2.5
		/// </summary>
		[CustomMethod]
		[SupportByVersion("ADODB", 2.5)]
        new void Resync();

		/// <summary>
		/// SupportByVersion ADODB 2.5
		/// </summary>
		/// <param name="affectRecords">optional NetOffice.ADODBApi.Enums.AffectEnum AffectRecords = 3</param>
		[CustomMethod]
		[SupportByVersion("ADODB", 2.5)]
        new void Resync(object affectRecords);

		/// <summary>
		/// SupportByVersion ADODB 2.5
		/// </summary>
		/// <param name="destination">optional object destination</param>
		/// <param name="persistFormat">optional NetOffice.ADODBApi.Enums.PersistFormatEnum PersistFormat = 0</param>
		[SupportByVersion("ADODB", 2.5)]
        new void Save(object destination, object persistFormat);

		/// <summary>
		/// SupportByVersion ADODB 2.5
		/// </summary>
		[CustomMethod]
		[SupportByVersion("ADODB", 2.5)]
        new void Save();

		/// <summary>
		/// SupportByVersion ADODB 2.5
		/// </summary>
		/// <param name="destination">optional object destination</param>
		[CustomMethod]
		[SupportByVersion("ADODB", 2.5)]
        new void Save(object destination);

		#endregion
	}
}
