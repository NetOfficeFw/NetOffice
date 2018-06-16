using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.DAOApi
{
	/// <summary>
	/// DispatchInterface Recordset 
	/// SupportByVersion DAO, 3.6,12.0
	/// </summary>
	[SupportByVersion("DAO", 3.6,12.0)]
	[EntityType(EntityType.IsDispatchInterface), BaseType]
	[TypeId("00000031-0000-0010-8000-00AA006D2EA4")]
	public interface Recordset : _DAO
	{
		#region Properties

		/// <summary>
		/// SupportByVersion DAO 3.6, 12.0
		/// Get
		/// </summary>
		[SupportByVersion("DAO", 3.6,12.0)]
		bool BOF { get; }

		/// <summary>
		/// SupportByVersion DAO 3.6, 12.0
		/// Get/Set
		/// </summary>
		[SupportByVersion("DAO", 3.6,12.0)]
		byte[] Bookmark { get; set; }

		/// <summary>
		/// SupportByVersion DAO 3.6, 12.0
		/// Get
		/// </summary>
		[SupportByVersion("DAO", 3.6,12.0)]
		bool Bookmarkable { get; }

		/// <summary>
		/// SupportByVersion DAO 3.6, 12.0
		/// Get
		/// </summary>
		[SupportByVersion("DAO", 3.6,12.0)]
		object DateCreated { get; }

		/// <summary>
		/// SupportByVersion DAO 3.6, 12.0
		/// Get
		/// </summary>
		[SupportByVersion("DAO", 3.6,12.0)]
		bool EOF { get; }

		/// <summary>
		/// SupportByVersion DAO 3.6, 12.0
		/// Get/Set
		/// </summary>
		[SupportByVersion("DAO", 3.6,12.0)]
		string Filter { get; set; }

		/// <summary>
		/// SupportByVersion DAO 3.6, 12.0
		/// Get/Set
		/// </summary>
		[SupportByVersion("DAO", 3.6,12.0)]
		string Index { get; set; }

		/// <summary>
		/// SupportByVersion DAO 3.6, 12.0
		/// Get
		/// </summary>
		[SupportByVersion("DAO", 3.6,12.0)]
		byte[] LastModified { get; }

		/// <summary>
		/// SupportByVersion DAO 3.6, 12.0
		/// Get
		/// </summary>
		[SupportByVersion("DAO", 3.6,12.0)]
		object LastUpdated { get; }

		/// <summary>
		/// SupportByVersion DAO 3.6, 12.0
		/// Get/Set
		/// </summary>
		[SupportByVersion("DAO", 3.6,12.0)]
		bool LockEdits { get; set; }

		/// <summary>
		/// SupportByVersion DAO 3.6, 12.0
		/// Get
		/// </summary>
		[SupportByVersion("DAO", 3.6,12.0)]
		string Name { get; }

		/// <summary>
		/// SupportByVersion DAO 3.6, 12.0
		/// Get
		/// </summary>
		[SupportByVersion("DAO", 3.6,12.0)]
		bool NoMatch { get; }

		/// <summary>
		/// SupportByVersion DAO 3.6, 12.0
		/// Get/Set
		/// </summary>
		[SupportByVersion("DAO", 3.6,12.0)]
		string Sort { get; set; }

		/// <summary>
		/// SupportByVersion DAO 3.6, 12.0
		/// Get
		/// </summary>
		[SupportByVersion("DAO", 3.6,12.0)]
		bool Transactions { get; }

		/// <summary>
		/// SupportByVersion DAO 3.6, 12.0
		/// Get
		/// </summary>
		[SupportByVersion("DAO", 3.6,12.0)]
		Int16 Type { get; }

		/// <summary>
		/// SupportByVersion DAO 3.6, 12.0
		/// Get
		/// </summary>
		[SupportByVersion("DAO", 3.6,12.0)]
		Int32 RecordCount { get; }

		/// <summary>
		/// SupportByVersion DAO 3.6, 12.0
		/// Get
		/// </summary>
		[SupportByVersion("DAO", 3.6,12.0)]
		bool Updatable { get; }

		/// <summary>
		/// SupportByVersion DAO 3.6, 12.0
		/// Get
		/// </summary>
		[SupportByVersion("DAO", 3.6,12.0)]
		bool Restartable { get; }

		/// <summary>
		/// SupportByVersion DAO 3.6, 12.0
		/// Get
		/// </summary>
		[SupportByVersion("DAO", 3.6,12.0)]
		string ValidationText { get; }

		/// <summary>
		/// SupportByVersion DAO 3.6, 12.0
		/// Get
		/// </summary>
		[SupportByVersion("DAO", 3.6,12.0)]
		string ValidationRule { get; }

		/// <summary>
		/// SupportByVersion DAO 3.6, 12.0
		/// Get/Set
		/// </summary>
		[SupportByVersion("DAO", 3.6,12.0)]
		byte[] CacheStart { get; set; }

		/// <summary>
		/// SupportByVersion DAO 3.6, 12.0
		/// Get/Set
		/// </summary>
		[SupportByVersion("DAO", 3.6,12.0)]
		Int32 CacheSize { get; set; }

		/// <summary>
		/// SupportByVersion DAO 3.6, 12.0
		/// Get/Set
		/// </summary>
		[SupportByVersion("DAO", 3.6,12.0)]
		Single PercentPosition { get; set; }

		/// <summary>
		/// SupportByVersion DAO 3.6, 12.0
		/// Get/Set
		/// </summary>
		[SupportByVersion("DAO", 3.6,12.0)]
		Int32 AbsolutePosition { get; set; }

		/// <summary>
		/// SupportByVersion DAO 3.6, 12.0
		/// Get
		/// </summary>
		[SupportByVersion("DAO", 3.6,12.0)]
		Int16 EditMode { get; }

		/// <summary>
		/// SupportByVersion DAO 3.6, 12.0
		/// Get
		/// </summary>
		[SupportByVersion("DAO", 3.6,12.0)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		Int32 ODBCFetchCount { get; }

		/// <summary>
		/// SupportByVersion DAO 3.6, 12.0
		/// Get
		/// </summary>
		[SupportByVersion("DAO", 3.6,12.0)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		Int32 ODBCFetchDelay { get; }

		/// <summary>
		/// SupportByVersion DAO 3.6, 12.0
		/// Get
		/// </summary>
		[SupportByVersion("DAO", 3.6,12.0)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		NetOffice.DAOApi.Database Parent { get; }

		/// <summary>
		/// SupportByVersion DAO 3.6, 12.0
		/// Get
		/// </summary>
		[SupportByVersion("DAO", 3.6,12.0)]
		NetOffice.DAOApi.Fields Fields { get; }

		/// <summary>
		/// SupportByVersion DAO 3.6, 12.0
		/// Get
		/// </summary>
		[SupportByVersion("DAO", 3.6,12.0)]
		NetOffice.DAOApi.Indexes Indexes { get; }

		/// <summary>
		/// SupportByVersion DAO 3.6, 12.0
		/// Get/Set
		/// </summary>
		/// <param name="item">object item</param>
		[SupportByVersion("DAO", 3.6,12.0)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		object get_Collect(object item);

        /// <summary>
        /// SupportByVersion DAO 3.6, 12.0
        /// Get/Set
        /// </summary>
        /// <param name="item">object item</param>
        ///  <param name="value">object value</param>
        [SupportByVersion("DAO", 3.6,12.0)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		void set_Collect(object item, object value);

		/// <summary>
		/// SupportByVersion DAO 3.6, 12.0
		/// Alias for get_Collect
		/// </summary>
		/// <param name="item">object item</param>
		[SupportByVersion("DAO", 3.6,12.0), Redirect("get_Collect")]
		object Collect(object item);

		/// <summary>
		/// SupportByVersion DAO 3.6, 12.0
		/// Get
		/// </summary>
		[SupportByVersion("DAO", 3.6,12.0)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		Int32 hStmt { get; }

		/// <summary>
		/// SupportByVersion DAO 3.6, 12.0
		/// Get
		/// </summary>
		[SupportByVersion("DAO", 3.6,12.0)]
		bool StillExecuting { get; }

		/// <summary>
		/// SupportByVersion DAO 3.6, 12.0
		/// Get/Set
		/// </summary>
		[SupportByVersion("DAO", 3.6,12.0)]
		Int32 BatchSize { get; set; }

		/// <summary>
		/// SupportByVersion DAO 3.6, 12.0
		/// Get
		/// </summary>
		[SupportByVersion("DAO", 3.6,12.0)]
		Int32 BatchCollisionCount { get; }

		/// <summary>
		/// SupportByVersion DAO 3.6, 12.0
		/// Get
		/// </summary>
		[SupportByVersion("DAO", 3.6,12.0)]
		object BatchCollisions { get; }

		/// <summary>
		/// SupportByVersion DAO 3.6, 12.0
		/// Get/Set
		/// </summary>
		[SupportByVersion("DAO", 3.6,12.0)]
		NetOffice.DAOApi.Connection Connection { get; set; }

		/// <summary>
		/// SupportByVersion DAO 3.6, 12.0
		/// Get
		/// </summary>
		[SupportByVersion("DAO", 3.6,12.0)]
		Int16 RecordStatus { get; }

		/// <summary>
		/// SupportByVersion DAO 3.6, 12.0
		/// Get/Set
		/// </summary>
		[SupportByVersion("DAO", 3.6,12.0)]
		Int32 UpdateOptions { get; set; }

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion DAO 3.6, 12.0
		/// </summary>
		[SupportByVersion("DAO", 3.6,12.0)]
		void _30_CancelUpdate();

		/// <summary>
		/// SupportByVersion DAO 3.6, 12.0
		/// </summary>
		[SupportByVersion("DAO", 3.6,12.0)]
		void AddNew();

		/// <summary>
		/// SupportByVersion DAO 3.6, 12.0
		/// </summary>
		[SupportByVersion("DAO", 3.6,12.0)]
		void Close();

		/// <summary>
		/// SupportByVersion DAO 3.6, 12.0
		/// </summary>
		/// <param name="type">optional object type</param>
		/// <param name="options">optional object options</param>
		[SupportByVersion("DAO", 3.6,12.0)]
		[BaseResult]
		NetOffice.DAOApi.Recordset OpenRecordset(object type, object options);

		/// <summary>
		/// SupportByVersion DAO 3.6, 12.0
		/// </summary>
		[CustomMethod]
		[BaseResult]
		[SupportByVersion("DAO", 3.6,12.0)]
		NetOffice.DAOApi.Recordset OpenRecordset();

		/// <summary>
		/// SupportByVersion DAO 3.6, 12.0
		/// </summary>
		/// <param name="type">optional object type</param>
		[CustomMethod]
		[BaseResult]
		[SupportByVersion("DAO", 3.6,12.0)]
		NetOffice.DAOApi.Recordset OpenRecordset(object type);

		/// <summary>
		/// SupportByVersion DAO 3.6, 12.0
		/// </summary>
		[SupportByVersion("DAO", 3.6,12.0)]
		void Delete();

		/// <summary>
		/// SupportByVersion DAO 3.6, 12.0
		/// </summary>
		[SupportByVersion("DAO", 3.6,12.0)]
		void Edit();

		/// <summary>
		/// SupportByVersion DAO 3.6, 12.0
		/// </summary>
		/// <param name="criteria">string criteria</param>
		[SupportByVersion("DAO", 3.6,12.0)]
		void FindFirst(string criteria);

		/// <summary>
		/// SupportByVersion DAO 3.6, 12.0
		/// </summary>
		/// <param name="criteria">string criteria</param>
		[SupportByVersion("DAO", 3.6,12.0)]
		void FindLast(string criteria);

		/// <summary>
		/// SupportByVersion DAO 3.6, 12.0
		/// </summary>
		/// <param name="criteria">string criteria</param>
		[SupportByVersion("DAO", 3.6,12.0)]
		void FindNext(string criteria);

		/// <summary>
		/// SupportByVersion DAO 3.6, 12.0
		/// </summary>
		/// <param name="criteria">string criteria</param>
		[SupportByVersion("DAO", 3.6,12.0)]
		void FindPrevious(string criteria);

		/// <summary>
		/// SupportByVersion DAO 3.6, 12.0
		/// </summary>
		[SupportByVersion("DAO", 3.6,12.0)]
		void MoveFirst();

		/// <summary>
		/// SupportByVersion DAO 3.6, 12.0
		/// </summary>
		[SupportByVersion("DAO", 3.6,12.0)]
		void _30_MoveLast();

		/// <summary>
		/// SupportByVersion DAO 3.6, 12.0
		/// </summary>
		[SupportByVersion("DAO", 3.6,12.0)]
		void MoveNext();

		/// <summary>
		/// SupportByVersion DAO 3.6, 12.0
		/// </summary>
		[SupportByVersion("DAO", 3.6,12.0)]
		void MovePrevious();

		/// <summary>
		/// SupportByVersion DAO 3.6, 12.0
		/// </summary>
		/// <param name="comparison">string comparison</param>
		/// <param name="key1">object key1</param>
		/// <param name="key2">optional object key2</param>
		/// <param name="key3">optional object key3</param>
		/// <param name="key4">optional object key4</param>
		/// <param name="key5">optional object key5</param>
		/// <param name="key6">optional object key6</param>
		/// <param name="key7">optional object key7</param>
		/// <param name="key8">optional object key8</param>
		/// <param name="key9">optional object key9</param>
		/// <param name="key10">optional object key10</param>
		/// <param name="key11">optional object key11</param>
		/// <param name="key12">optional object key12</param>
		/// <param name="key13">optional object key13</param>
		[SupportByVersion("DAO", 3.6,12.0)]
		void Seek(string comparison, object key1, object key2, object key3, object key4, object key5, object key6, object key7, object key8, object key9, object key10, object key11, object key12, object key13);

		/// <summary>
		/// SupportByVersion DAO 3.6, 12.0
		/// </summary>
		/// <param name="comparison">string comparison</param>
		/// <param name="key1">object key1</param>
		[CustomMethod]
		[SupportByVersion("DAO", 3.6,12.0)]
		void Seek(string comparison, object key1);

		/// <summary>
		/// SupportByVersion DAO 3.6, 12.0
		/// </summary>
		/// <param name="comparison">string comparison</param>
		/// <param name="key1">object key1</param>
		/// <param name="key2">optional object key2</param>
		[CustomMethod]
		[SupportByVersion("DAO", 3.6,12.0)]
		void Seek(string comparison, object key1, object key2);

		/// <summary>
		/// SupportByVersion DAO 3.6, 12.0
		/// </summary>
		/// <param name="comparison">string comparison</param>
		/// <param name="key1">object key1</param>
		/// <param name="key2">optional object key2</param>
		/// <param name="key3">optional object key3</param>
		[CustomMethod]
		[SupportByVersion("DAO", 3.6,12.0)]
		void Seek(string comparison, object key1, object key2, object key3);

		/// <summary>
		/// SupportByVersion DAO 3.6, 12.0
		/// </summary>
		/// <param name="comparison">string comparison</param>
		/// <param name="key1">object key1</param>
		/// <param name="key2">optional object key2</param>
		/// <param name="key3">optional object key3</param>
		/// <param name="key4">optional object key4</param>
		[CustomMethod]
		[SupportByVersion("DAO", 3.6,12.0)]
		void Seek(string comparison, object key1, object key2, object key3, object key4);

		/// <summary>
		/// SupportByVersion DAO 3.6, 12.0
		/// </summary>
		/// <param name="comparison">string comparison</param>
		/// <param name="key1">object key1</param>
		/// <param name="key2">optional object key2</param>
		/// <param name="key3">optional object key3</param>
		/// <param name="key4">optional object key4</param>
		/// <param name="key5">optional object key5</param>
		[CustomMethod]
		[SupportByVersion("DAO", 3.6,12.0)]
		void Seek(string comparison, object key1, object key2, object key3, object key4, object key5);

		/// <summary>
		/// SupportByVersion DAO 3.6, 12.0
		/// </summary>
		/// <param name="comparison">string comparison</param>
		/// <param name="key1">object key1</param>
		/// <param name="key2">optional object key2</param>
		/// <param name="key3">optional object key3</param>
		/// <param name="key4">optional object key4</param>
		/// <param name="key5">optional object key5</param>
		/// <param name="key6">optional object key6</param>
		[CustomMethod]
		[SupportByVersion("DAO", 3.6,12.0)]
		void Seek(string comparison, object key1, object key2, object key3, object key4, object key5, object key6);

		/// <summary>
		/// SupportByVersion DAO 3.6, 12.0
		/// </summary>
		/// <param name="comparison">string comparison</param>
		/// <param name="key1">object key1</param>
		/// <param name="key2">optional object key2</param>
		/// <param name="key3">optional object key3</param>
		/// <param name="key4">optional object key4</param>
		/// <param name="key5">optional object key5</param>
		/// <param name="key6">optional object key6</param>
		/// <param name="key7">optional object key7</param>
		[CustomMethod]
		[SupportByVersion("DAO", 3.6,12.0)]
		void Seek(string comparison, object key1, object key2, object key3, object key4, object key5, object key6, object key7);

		/// <summary>
		/// SupportByVersion DAO 3.6, 12.0
		/// </summary>
		/// <param name="comparison">string comparison</param>
		/// <param name="key1">object key1</param>
		/// <param name="key2">optional object key2</param>
		/// <param name="key3">optional object key3</param>
		/// <param name="key4">optional object key4</param>
		/// <param name="key5">optional object key5</param>
		/// <param name="key6">optional object key6</param>
		/// <param name="key7">optional object key7</param>
		/// <param name="key8">optional object key8</param>
		[CustomMethod]
		[SupportByVersion("DAO", 3.6,12.0)]
		void Seek(string comparison, object key1, object key2, object key3, object key4, object key5, object key6, object key7, object key8);

		/// <summary>
		/// SupportByVersion DAO 3.6, 12.0
		/// </summary>
		/// <param name="comparison">string comparison</param>
		/// <param name="key1">object key1</param>
		/// <param name="key2">optional object key2</param>
		/// <param name="key3">optional object key3</param>
		/// <param name="key4">optional object key4</param>
		/// <param name="key5">optional object key5</param>
		/// <param name="key6">optional object key6</param>
		/// <param name="key7">optional object key7</param>
		/// <param name="key8">optional object key8</param>
		/// <param name="key9">optional object key9</param>
		[CustomMethod]
		[SupportByVersion("DAO", 3.6,12.0)]
		void Seek(string comparison, object key1, object key2, object key3, object key4, object key5, object key6, object key7, object key8, object key9);

		/// <summary>
		/// SupportByVersion DAO 3.6, 12.0
		/// </summary>
		/// <param name="comparison">string comparison</param>
		/// <param name="key1">object key1</param>
		/// <param name="key2">optional object key2</param>
		/// <param name="key3">optional object key3</param>
		/// <param name="key4">optional object key4</param>
		/// <param name="key5">optional object key5</param>
		/// <param name="key6">optional object key6</param>
		/// <param name="key7">optional object key7</param>
		/// <param name="key8">optional object key8</param>
		/// <param name="key9">optional object key9</param>
		/// <param name="key10">optional object key10</param>
		[CustomMethod]
		[SupportByVersion("DAO", 3.6,12.0)]
		void Seek(string comparison, object key1, object key2, object key3, object key4, object key5, object key6, object key7, object key8, object key9, object key10);

		/// <summary>
		/// SupportByVersion DAO 3.6, 12.0
		/// </summary>
		/// <param name="comparison">string comparison</param>
		/// <param name="key1">object key1</param>
		/// <param name="key2">optional object key2</param>
		/// <param name="key3">optional object key3</param>
		/// <param name="key4">optional object key4</param>
		/// <param name="key5">optional object key5</param>
		/// <param name="key6">optional object key6</param>
		/// <param name="key7">optional object key7</param>
		/// <param name="key8">optional object key8</param>
		/// <param name="key9">optional object key9</param>
		/// <param name="key10">optional object key10</param>
		/// <param name="key11">optional object key11</param>
		[CustomMethod]
		[SupportByVersion("DAO", 3.6,12.0)]
		void Seek(string comparison, object key1, object key2, object key3, object key4, object key5, object key6, object key7, object key8, object key9, object key10, object key11);

		/// <summary>
		/// SupportByVersion DAO 3.6, 12.0
		/// </summary>
		/// <param name="comparison">string comparison</param>
		/// <param name="key1">object key1</param>
		/// <param name="key2">optional object key2</param>
		/// <param name="key3">optional object key3</param>
		/// <param name="key4">optional object key4</param>
		/// <param name="key5">optional object key5</param>
		/// <param name="key6">optional object key6</param>
		/// <param name="key7">optional object key7</param>
		/// <param name="key8">optional object key8</param>
		/// <param name="key9">optional object key9</param>
		/// <param name="key10">optional object key10</param>
		/// <param name="key11">optional object key11</param>
		/// <param name="key12">optional object key12</param>
		[CustomMethod]
		[SupportByVersion("DAO", 3.6,12.0)]
		void Seek(string comparison, object key1, object key2, object key3, object key4, object key5, object key6, object key7, object key8, object key9, object key10, object key11, object key12);

		/// <summary>
		/// SupportByVersion DAO 3.6, 12.0
		/// </summary>
		[SupportByVersion("DAO", 3.6,12.0)]
		void _30_Update();

		/// <summary>
		/// SupportByVersion DAO 3.6, 12.0
		/// </summary>
		[SupportByVersion("DAO", 3.6,12.0)]
		[BaseResult]
		new NetOffice.DAOApi.Recordset Clone();

		/// <summary>
		/// SupportByVersion DAO 3.6, 12.0
		/// </summary>
		/// <param name="newQueryDef">optional object newQueryDef</param>
		[SupportByVersion("DAO", 3.6,12.0)]
		void Requery(object newQueryDef);

		/// <summary>
		/// SupportByVersion DAO 3.6, 12.0
		/// </summary>
		[CustomMethod]
		[SupportByVersion("DAO", 3.6,12.0)]
		void Requery();

		/// <summary>
		/// SupportByVersion DAO 3.6, 12.0
		/// </summary>
		/// <param name="rows">Int32 rows</param>
		/// <param name="startBookmark">optional object startBookmark</param>
		[SupportByVersion("DAO", 3.6,12.0)]
		void Move(Int32 rows, object startBookmark);

		/// <summary>
		/// SupportByVersion DAO 3.6, 12.0
		/// </summary>
		/// <param name="rows">Int32 rows</param>
		[CustomMethod]
		[SupportByVersion("DAO", 3.6,12.0)]
		void Move(Int32 rows);

		/// <summary>
		/// SupportByVersion DAO 3.6, 12.0
		/// </summary>
		/// <param name="rows">optional object rows</param>
		/// <param name="startBookmark">optional object startBookmark</param>
		[SupportByVersion("DAO", 3.6,12.0)]
		void FillCache(object rows, object startBookmark);

		/// <summary>
		/// SupportByVersion DAO 3.6, 12.0
		/// </summary>
		[CustomMethod]
		[SupportByVersion("DAO", 3.6,12.0)]
		void FillCache();

		/// <summary>
		/// SupportByVersion DAO 3.6, 12.0
		/// </summary>
		/// <param name="rows">optional object rows</param>
		[CustomMethod]
		[SupportByVersion("DAO", 3.6,12.0)]
		void FillCache(object rows);

		/// <summary>
		/// SupportByVersion DAO 3.6, 12.0
		/// </summary>
		/// <param name="options">optional object options</param>
		/// <param name="inconsistent">optional object inconsistent</param>
		[SupportByVersion("DAO", 3.6,12.0)]
		[BaseResult]
		NetOffice.DAOApi.Recordset CreateDynaset(object options, object inconsistent);

		/// <summary>
		/// SupportByVersion DAO 3.6, 12.0
		/// </summary>
		[CustomMethod]
		[BaseResult]
		[SupportByVersion("DAO", 3.6,12.0)]
		NetOffice.DAOApi.Recordset CreateDynaset();

		/// <summary>
		/// SupportByVersion DAO 3.6, 12.0
		/// </summary>
		/// <param name="options">optional object options</param>
		[CustomMethod]
		[BaseResult]
		[SupportByVersion("DAO", 3.6,12.0)]
		NetOffice.DAOApi.Recordset CreateDynaset(object options);

		/// <summary>
		/// SupportByVersion DAO 3.6, 12.0
		/// </summary>
		/// <param name="options">optional object options</param>
		[SupportByVersion("DAO", 3.6,12.0)]
		[BaseResult]
		NetOffice.DAOApi.Recordset CreateSnapshot(object options);

		/// <summary>
		/// SupportByVersion DAO 3.6, 12.0
		/// </summary>
		[CustomMethod]
		[BaseResult]
		[SupportByVersion("DAO", 3.6,12.0)]
		NetOffice.DAOApi.Recordset CreateSnapshot();

		/// <summary>
		/// SupportByVersion DAO 3.6, 12.0
		/// </summary>
		[SupportByVersion("DAO", 3.6,12.0)]
		NetOffice.DAOApi.QueryDef CopyQueryDef();

		/// <summary>
		/// SupportByVersion DAO 3.6, 12.0
		/// </summary>
		[SupportByVersion("DAO", 3.6,12.0)]
		[BaseResult]
		NetOffice.DAOApi.Recordset ListFields();

		/// <summary>
		/// SupportByVersion DAO 3.6, 12.0
		/// </summary>
		[SupportByVersion("DAO", 3.6,12.0)]
		[BaseResult]
		NetOffice.DAOApi.Recordset ListIndexes();

		/// <summary>
		/// SupportByVersion DAO 3.6, 12.0
		/// </summary>
		/// <param name="numRows">optional object numRows</param>
		[SupportByVersion("DAO", 3.6,12.0)]
		object GetRows(object numRows);

		/// <summary>
		/// SupportByVersion DAO 3.6, 12.0
		/// </summary>
		[CustomMethod]
		[SupportByVersion("DAO", 3.6,12.0)]
		object GetRows();

		/// <summary>
		/// SupportByVersion DAO 3.6, 12.0
		/// </summary>
		[SupportByVersion("DAO", 3.6,12.0)]
		void Cancel();

		/// <summary>
		/// SupportByVersion DAO 3.6, 12.0
		/// </summary>
		[SupportByVersion("DAO", 3.6,12.0)]
		bool NextRecordset();

		/// <summary>
		/// SupportByVersion DAO 3.6, 12.0
		/// </summary>
		/// <param name="updateType">optional Int32 UpdateType = 1</param>
		[SupportByVersion("DAO", 3.6,12.0)]
		void CancelUpdate(object updateType);

		/// <summary>
		/// SupportByVersion DAO 3.6, 12.0
		/// </summary>
		[CustomMethod]
		[SupportByVersion("DAO", 3.6,12.0)]
		void CancelUpdate();

		/// <summary>
		/// SupportByVersion DAO 3.6, 12.0
		/// </summary>
		/// <param name="updateType">optional Int32 UpdateType = 1</param>
		/// <param name="force">optional bool Force = false</param>
		[SupportByVersion("DAO", 3.6,12.0)]
		void Update(object updateType, object force);

		/// <summary>
		/// SupportByVersion DAO 3.6, 12.0
		/// </summary>
		[CustomMethod]
		[SupportByVersion("DAO", 3.6,12.0)]
		void Update();

		/// <summary>
		/// SupportByVersion DAO 3.6, 12.0
		/// </summary>
		/// <param name="updateType">optional Int32 UpdateType = 1</param>
		[CustomMethod]
		[SupportByVersion("DAO", 3.6,12.0)]
		void Update(object updateType);

		/// <summary>
		/// SupportByVersion DAO 3.6, 12.0
		/// </summary>
		/// <param name="options">optional Int32 Options = 0</param>
		[SupportByVersion("DAO", 3.6,12.0)]
		void MoveLast(object options);

		/// <summary>
		/// SupportByVersion DAO 3.6, 12.0
		/// </summary>
		[CustomMethod]
		[SupportByVersion("DAO", 3.6,12.0)]
		void MoveLast();

		#endregion
	}
}
