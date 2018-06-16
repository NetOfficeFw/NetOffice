using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.ADODBApi
{
	/// <summary>
	/// DispatchInterface Recordset15_Deprecated 
	/// SupportByVersion ADODB, 2.5
	/// </summary>
	[SupportByVersion("ADODB", 2.5)]
	[EntityType(EntityType.IsDispatchInterface), BaseType]
	[TypeId("0000050E-0000-0010-8000-00AA006D2EA4")]
	public interface Recordset15_Deprecated : _ADO
	{
		#region Properties

		/// <summary>
		/// SupportByVersion ADODB 2.5
		/// Get/Set
		/// </summary>
		[SupportByVersion("ADODB", 2.5)]
		NetOffice.ADODBApi.Enums.PositionEnum AbsolutePosition { get; set; }

		/// <summary>
		/// SupportByVersion ADODB 2.5
		/// Get/Set
		/// Unknown COM Proxy
		/// </summary>
		[SupportByVersion("ADODB", 2.5), ProxyResult]
		object ActiveConnection { get; set; }

		/// <summary>
		/// SupportByVersion ADODB 2.5
		/// Get
		/// </summary>
		[SupportByVersion("ADODB", 2.5)]
		bool BOF { get; }

		/// <summary>
		/// SupportByVersion ADODB 2.5
		/// Get/Set
		/// </summary>
		[SupportByVersion("ADODB", 2.5)]
		object Bookmark { get; set; }

		/// <summary>
		/// SupportByVersion ADODB 2.5
		/// Get/Set
		/// </summary>
		[SupportByVersion("ADODB", 2.5)]
		Int32 CacheSize { get; set; }

		/// <summary>
		/// SupportByVersion ADODB 2.5
		/// Get/Set
		/// </summary>
		[SupportByVersion("ADODB", 2.5)]
		NetOffice.ADODBApi.Enums.CursorTypeEnum CursorType { get; set; }

		/// <summary>
		/// SupportByVersion ADODB 2.5
		/// Get
		/// </summary>
		[SupportByVersion("ADODB", 2.5)]
		bool EOF { get; }

		/// <summary>
		/// SupportByVersion ADODB 2.5
		/// Get
		/// </summary>
		[SupportByVersion("ADODB", 2.5)]
		NetOffice.ADODBApi.Fields_Deprecated Fields { get; }

		/// <summary>
		/// SupportByVersion ADODB 2.5
		/// Get/Set
		/// </summary>
		[SupportByVersion("ADODB", 2.5)]
		NetOffice.ADODBApi.Enums.LockTypeEnum LockType { get; set; }

		/// <summary>
		/// SupportByVersion ADODB 2.5
		/// Get/Set
		/// </summary>
		[SupportByVersion("ADODB", 2.5)]
		Int32 MaxRecords { get; set; }

		/// <summary>
		/// SupportByVersion ADODB 2.5
		/// Get
		/// </summary>
		[SupportByVersion("ADODB", 2.5)]
		Int32 RecordCount { get; }

		/// <summary>
		/// SupportByVersion ADODB 2.5
		/// Get/Set
		/// Unknown COM Proxy
		/// </summary>
		[SupportByVersion("ADODB", 2.5), ProxyResult]
		object Source { get; set; }

		/// <summary>
		/// SupportByVersion ADODB 2.5
		/// Get/Set
		/// </summary>
		[SupportByVersion("ADODB", 2.5)]
		NetOffice.ADODBApi.Enums.PositionEnum AbsolutePage { get; set; }

		/// <summary>
		/// SupportByVersion ADODB 2.5
		/// Get
		/// </summary>
		[SupportByVersion("ADODB", 2.5)]
		NetOffice.ADODBApi.Enums.EditModeEnum EditMode { get; }

		/// <summary>
		/// SupportByVersion ADODB 2.5
		/// Get/Set
		/// </summary>
		[SupportByVersion("ADODB", 2.5)]
		object Filter { get; set; }

		/// <summary>
		/// SupportByVersion ADODB 2.5
		/// Get
		/// </summary>
		[SupportByVersion("ADODB", 2.5)]
		Int32 PageCount { get; }

		/// <summary>
		/// SupportByVersion ADODB 2.5
		/// Get/Set
		/// </summary>
		[SupportByVersion("ADODB", 2.5)]
		Int32 PageSize { get; set; }

		/// <summary>
		/// SupportByVersion ADODB 2.5
		/// Get/Set
		/// </summary>
		[SupportByVersion("ADODB", 2.5)]
		string Sort { get; set; }

		/// <summary>
		/// SupportByVersion ADODB 2.5
		/// Get
		/// </summary>
		[SupportByVersion("ADODB", 2.5)]
		Int32 Status { get; }

		/// <summary>
		/// SupportByVersion ADODB 2.5
		/// Get
		/// </summary>
		[SupportByVersion("ADODB", 2.5)]
		Int32 State { get; }

		/// <summary>
		/// SupportByVersion ADODB 2.5
		/// Get/Set
		/// </summary>
		[SupportByVersion("ADODB", 2.5)]
		NetOffice.ADODBApi.Enums.CursorLocationEnum CursorLocation { get; set; }

		/// <summary>
		/// SupportByVersion ADODB 2.5
		/// Get/Set
		/// </summary>
		/// <param name="index">object index</param>
		[SupportByVersion("ADODB", 2.5)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		object get_Collect(object index);

        /// <summary>
        /// SupportByVersion ADODB 2.5
        /// Get/Set
        /// </summary>
        /// <param name="index">object index</param>
        ///  <param name="value">object value</param>
        [SupportByVersion("ADODB", 2.5)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		void set_Collect(object index, object value);

		/// <summary>
		/// SupportByVersion ADODB 2.5
		/// Alias for get_Collect
		/// </summary>
		/// <param name="index">object index</param>
		[SupportByVersion("ADODB", 2.5), Redirect("get_Collect")]
		object Collect(object index);

		/// <summary>
		/// SupportByVersion ADODB 2.5
		/// Get/Set
		/// </summary>
		[SupportByVersion("ADODB", 2.5)]
		NetOffice.ADODBApi.Enums.MarshalOptionsEnum MarshalOptions { get; set; }

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion ADODB 2.5
		/// </summary>
		/// <param name="fieldList">optional object fieldList</param>
		/// <param name="values">optional object values</param>
		[SupportByVersion("ADODB", 2.5)]
		void AddNew(object fieldList, object values);

		/// <summary>
		/// SupportByVersion ADODB 2.5
		/// </summary>
		[CustomMethod]
		[SupportByVersion("ADODB", 2.5)]
		void AddNew();

		/// <summary>
		/// SupportByVersion ADODB 2.5
		/// </summary>
		/// <param name="fieldList">optional object fieldList</param>
		[CustomMethod]
		[SupportByVersion("ADODB", 2.5)]
		void AddNew(object fieldList);

		/// <summary>
		/// SupportByVersion ADODB 2.5
		/// </summary>
		[SupportByVersion("ADODB", 2.5)]
		void CancelUpdate();

		/// <summary>
		/// SupportByVersion ADODB 2.5
		/// </summary>
		[SupportByVersion("ADODB", 2.5)]
		void Close();

		/// <summary>
		/// SupportByVersion ADODB 2.5
		/// </summary>
		/// <param name="affectRecords">optional NetOffice.ADODBApi.Enums.AffectEnum AffectRecords = 1</param>
		[SupportByVersion("ADODB", 2.5)]
		void Delete(object affectRecords);

		/// <summary>
		/// SupportByVersion ADODB 2.5
		/// </summary>
		[CustomMethod]
		[SupportByVersion("ADODB", 2.5)]
		void Delete();

		/// <summary>
		/// SupportByVersion ADODB 2.5
		/// </summary>
		/// <param name="rows">optional Int32 Rows = -1</param>
		/// <param name="start">optional object start</param>
		/// <param name="fields">optional object fields</param>
		[SupportByVersion("ADODB", 2.5)]
		object GetRows(object rows, object start, object fields);

		/// <summary>
		/// SupportByVersion ADODB 2.5
		/// </summary>
		[CustomMethod]
		[SupportByVersion("ADODB", 2.5)]
		object GetRows();

		/// <summary>
		/// SupportByVersion ADODB 2.5
		/// </summary>
		/// <param name="rows">optional Int32 Rows = -1</param>
		[CustomMethod]
		[SupportByVersion("ADODB", 2.5)]
		object GetRows(object rows);

		/// <summary>
		/// SupportByVersion ADODB 2.5
		/// </summary>
		/// <param name="rows">optional Int32 Rows = -1</param>
		/// <param name="start">optional object start</param>
		[CustomMethod]
		[SupportByVersion("ADODB", 2.5)]
		object GetRows(object rows, object start);

		/// <summary>
		/// SupportByVersion ADODB 2.5
		/// </summary>
		/// <param name="numRecords">Int32 numRecords</param>
		/// <param name="start">optional object start</param>
		[SupportByVersion("ADODB", 2.5)]
		void Move(Int32 numRecords, object start);

		/// <summary>
		/// SupportByVersion ADODB 2.5
		/// </summary>
		/// <param name="numRecords">Int32 numRecords</param>
		[CustomMethod]
		[SupportByVersion("ADODB", 2.5)]
		void Move(Int32 numRecords);

		/// <summary>
		/// SupportByVersion ADODB 2.5
		/// </summary>
		[SupportByVersion("ADODB", 2.5)]
		void MoveNext();

		/// <summary>
		/// SupportByVersion ADODB 2.5
		/// </summary>
		[SupportByVersion("ADODB", 2.5)]
		void MovePrevious();

		/// <summary>
		/// SupportByVersion ADODB 2.5
		/// </summary>
		[SupportByVersion("ADODB", 2.5)]
		void MoveFirst();

		/// <summary>
		/// SupportByVersion ADODB 2.5
		/// </summary>
		[SupportByVersion("ADODB", 2.5)]
		void MoveLast();

		/// <summary>
		/// SupportByVersion ADODB 2.5
		/// </summary>
		/// <param name="source">optional object source</param>
		/// <param name="activeConnection">optional object activeConnection</param>
		/// <param name="cursorType">optional NetOffice.ADODBApi.Enums.CursorTypeEnum CursorType = -1</param>
		/// <param name="lockType">optional NetOffice.ADODBApi.Enums.LockTypeEnum LockType = -1</param>
		/// <param name="options">optional Int32 Options = -1</param>
		[SupportByVersion("ADODB", 2.5)]
		void Open(object source, object activeConnection, object cursorType, object lockType, object options);

		/// <summary>
		/// SupportByVersion ADODB 2.5
		/// </summary>
		[CustomMethod]
		[SupportByVersion("ADODB", 2.5)]
		void Open();

		/// <summary>
		/// SupportByVersion ADODB 2.5
		/// </summary>
		/// <param name="source">optional object source</param>
		[CustomMethod]
		[SupportByVersion("ADODB", 2.5)]
		void Open(object source);

		/// <summary>
		/// SupportByVersion ADODB 2.5
		/// </summary>
		/// <param name="source">optional object source</param>
		/// <param name="activeConnection">optional object activeConnection</param>
		[CustomMethod]
		[SupportByVersion("ADODB", 2.5)]
		void Open(object source, object activeConnection);

		/// <summary>
		/// SupportByVersion ADODB 2.5
		/// </summary>
		/// <param name="source">optional object source</param>
		/// <param name="activeConnection">optional object activeConnection</param>
		/// <param name="cursorType">optional NetOffice.ADODBApi.Enums.CursorTypeEnum CursorType = -1</param>
		[CustomMethod]
		[SupportByVersion("ADODB", 2.5)]
		void Open(object source, object activeConnection, object cursorType);

		/// <summary>
		/// SupportByVersion ADODB 2.5
		/// </summary>
		/// <param name="source">optional object source</param>
		/// <param name="activeConnection">optional object activeConnection</param>
		/// <param name="cursorType">optional NetOffice.ADODBApi.Enums.CursorTypeEnum CursorType = -1</param>
		/// <param name="lockType">optional NetOffice.ADODBApi.Enums.LockTypeEnum LockType = -1</param>
		[CustomMethod]
		[SupportByVersion("ADODB", 2.5)]
		void Open(object source, object activeConnection, object cursorType, object lockType);

		/// <summary>
		/// SupportByVersion ADODB 2.5
		/// </summary>
		/// <param name="options">optional Int32 Options = -1</param>
		[SupportByVersion("ADODB", 2.5)]
		void Requery(object options);

		/// <summary>
		/// SupportByVersion ADODB 2.5
		/// </summary>
		[CustomMethod]
		[SupportByVersion("ADODB", 2.5)]
		void Requery();

		/// <summary>
		/// SupportByVersion ADODB 2.5
		/// </summary>
		/// <param name="affectRecords">optional NetOffice.ADODBApi.Enums.AffectEnum AffectRecords = 3</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("ADODB", 2.5)]
		void _xResync(object affectRecords);

		/// <summary>
		/// SupportByVersion ADODB 2.5
		/// </summary>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("ADODB", 2.5)]
		void _xResync();

		/// <summary>
		/// SupportByVersion ADODB 2.5
		/// </summary>
		/// <param name="fields">optional object fields</param>
		/// <param name="values">optional object values</param>
		[SupportByVersion("ADODB", 2.5)]
		void Update(object fields, object values);

		/// <summary>
		/// SupportByVersion ADODB 2.5
		/// </summary>
		[CustomMethod]
		[SupportByVersion("ADODB", 2.5)]
		void Update();

		/// <summary>
		/// SupportByVersion ADODB 2.5
		/// </summary>
		/// <param name="fields">optional object fields</param>
		[CustomMethod]
		[SupportByVersion("ADODB", 2.5)]
		void Update(object fields);

		/// <summary>
		/// SupportByVersion ADODB 2.5
		/// </summary>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("ADODB", 2.5)]
		NetOffice.ADODBApi._Recordset_Deprecated _xClone();

		/// <summary>
		/// SupportByVersion ADODB 2.5
		/// </summary>
		/// <param name="affectRecords">optional NetOffice.ADODBApi.Enums.AffectEnum AffectRecords = 3</param>
		[SupportByVersion("ADODB", 2.5)]
		void UpdateBatch(object affectRecords);

		/// <summary>
		/// SupportByVersion ADODB 2.5
		/// </summary>
		[CustomMethod]
		[SupportByVersion("ADODB", 2.5)]
		void UpdateBatch();

		/// <summary>
		/// SupportByVersion ADODB 2.5
		/// </summary>
		/// <param name="affectRecords">optional NetOffice.ADODBApi.Enums.AffectEnum AffectRecords = 3</param>
		[SupportByVersion("ADODB", 2.5)]
		void CancelBatch(object affectRecords);

		/// <summary>
		/// SupportByVersion ADODB 2.5
		/// </summary>
		[CustomMethod]
		[SupportByVersion("ADODB", 2.5)]
		void CancelBatch();

		/// <summary>
		/// SupportByVersion ADODB 2.5
		/// </summary>
		/// <param name="recordsAffected">optional object recordsAffected</param>
		[SupportByVersion("ADODB", 2.5)]
		NetOffice.ADODBApi._Recordset_Deprecated NextRecordset(object recordsAffected);

		/// <summary>
		/// SupportByVersion ADODB 2.5
		/// </summary>
		[CustomMethod]
		[SupportByVersion("ADODB", 2.5)]
		NetOffice.ADODBApi._Recordset_Deprecated NextRecordset();

		/// <summary>
		/// SupportByVersion ADODB 2.5
		/// </summary>
		/// <param name="cursorOptions">NetOffice.ADODBApi.Enums.CursorOptionEnum cursorOptions</param>
		[SupportByVersion("ADODB", 2.5)]
		bool Supports(NetOffice.ADODBApi.Enums.CursorOptionEnum cursorOptions);

		/// <summary>
		/// SupportByVersion ADODB 2.5
		/// </summary>
		/// <param name="criteria">string criteria</param>
		/// <param name="skipRecords">optional Int32 SkipRecords = 0</param>
		/// <param name="searchDirection">optional NetOffice.ADODBApi.Enums.SearchDirectionEnum SearchDirection = 1</param>
		/// <param name="start">optional object start</param>
		[SupportByVersion("ADODB", 2.5)]
		void Find(string criteria, object skipRecords, object searchDirection, object start);

		/// <summary>
		/// SupportByVersion ADODB 2.5
		/// </summary>
		/// <param name="criteria">string criteria</param>
		[CustomMethod]
		[SupportByVersion("ADODB", 2.5)]
		void Find(string criteria);

		/// <summary>
		/// SupportByVersion ADODB 2.5
		/// </summary>
		/// <param name="criteria">string criteria</param>
		/// <param name="skipRecords">optional Int32 SkipRecords = 0</param>
		[CustomMethod]
		[SupportByVersion("ADODB", 2.5)]
		void Find(string criteria, object skipRecords);

		/// <summary>
		/// SupportByVersion ADODB 2.5
		/// </summary>
		/// <param name="criteria">string criteria</param>
		/// <param name="skipRecords">optional Int32 SkipRecords = 0</param>
		/// <param name="searchDirection">optional NetOffice.ADODBApi.Enums.SearchDirectionEnum SearchDirection = 1</param>
		[CustomMethod]
		[SupportByVersion("ADODB", 2.5)]
		void Find(string criteria, object skipRecords, object searchDirection);

		#endregion
	}
}
