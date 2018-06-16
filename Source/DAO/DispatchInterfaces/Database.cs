using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.DAOApi
{
	/// <summary>
	/// DispatchInterface Database 
	/// SupportByVersion DAO, 3.6,12.0
	/// </summary>
	[SupportByVersion("DAO", 3.6,12.0)]
	[EntityType(EntityType.IsDispatchInterface)]
	[TypeId("00000071-0000-0010-8000-00AA006D2EA4")]
	public interface Database : _DAO
	{
		#region Properties

		/// <summary>
		/// SupportByVersion DAO 3.6, 12.0
		/// Get
		/// </summary>
		[SupportByVersion("DAO", 3.6,12.0)]
		Int32 CollatingOrder { get; }

		/// <summary>
		/// SupportByVersion DAO 3.6, 12.0
		/// Get/Set
		/// </summary>
		[SupportByVersion("DAO", 3.6,12.0)]
		string Connect { get; set; }

		/// <summary>
		/// SupportByVersion DAO 3.6, 12.0
		/// Get
		/// </summary>
		[SupportByVersion("DAO", 3.6,12.0)]
		string Name { get; }

		/// <summary>
		/// SupportByVersion DAO 3.6, 12.0
		/// Get/Set
		/// </summary>
		[SupportByVersion("DAO", 3.6,12.0)]
		Int16 QueryTimeout { get; set; }

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
		bool Updatable { get; }

		/// <summary>
		/// SupportByVersion DAO 3.6, 12.0
		/// Get
		/// </summary>
		[SupportByVersion("DAO", 3.6,12.0)]
		string Version { get; }

		/// <summary>
		/// SupportByVersion DAO 3.6, 12.0
		/// Get
		/// </summary>
		[SupportByVersion("DAO", 3.6,12.0)]
		Int32 RecordsAffected { get; }

		/// <summary>
		/// SupportByVersion DAO 3.6, 12.0
		/// Get
		/// </summary>
		[SupportByVersion("DAO", 3.6,12.0)]
		NetOffice.DAOApi.TableDefs TableDefs { get; }

		/// <summary>
		/// SupportByVersion DAO 3.6, 12.0
		/// Get
		/// </summary>
		[SupportByVersion("DAO", 3.6,12.0)]
		NetOffice.DAOApi.QueryDefs QueryDefs { get; }

		/// <summary>
		/// SupportByVersion DAO 3.6, 12.0
		/// Get
		/// </summary>
		[SupportByVersion("DAO", 3.6,12.0)]
		NetOffice.DAOApi.Relations Relations { get; }

		/// <summary>
		/// SupportByVersion DAO 3.6, 12.0
		/// Get
		/// </summary>
		[SupportByVersion("DAO", 3.6,12.0)]
		NetOffice.DAOApi.Containers Containers { get; }

		/// <summary>
		/// SupportByVersion DAO 3.6, 12.0
		/// Get
		/// </summary>
		[SupportByVersion("DAO", 3.6,12.0)]
		NetOffice.DAOApi.Recordsets Recordsets { get; }

		/// <summary>
		/// SupportByVersion DAO 3.6, 12.0
		/// Get
		/// </summary>
		[SupportByVersion("DAO", 3.6,12.0)]
		string ReplicaID { get; }

		/// <summary>
		/// SupportByVersion DAO 3.6, 12.0
		/// Get/Set
		/// </summary>
		[SupportByVersion("DAO", 3.6,12.0)]
		string DesignMasterID { get; set; }

		/// <summary>
		/// SupportByVersion DAO 3.6, 12.0
		/// Get
		/// </summary>
		[SupportByVersion("DAO", 3.6,12.0)]
		NetOffice.DAOApi.Connection Connection { get; }

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion DAO 3.6, 12.0
		/// </summary>
		[SupportByVersion("DAO", 3.6,12.0)]
		void Close();

		/// <summary>
		/// SupportByVersion DAO 3.6, 12.0
		/// </summary>
		/// <param name="query">string query</param>
		/// <param name="options">optional object options</param>
		[SupportByVersion("DAO", 3.6,12.0)]
		void Execute(string query, object options);

		/// <summary>
		/// SupportByVersion DAO 3.6, 12.0
		/// </summary>
		/// <param name="query">string query</param>
		[CustomMethod]
		[SupportByVersion("DAO", 3.6,12.0)]
		void Execute(string query);

		/// <summary>
		/// SupportByVersion DAO 3.6, 12.0
		/// </summary>
		/// <param name="name">string name</param>
		/// <param name="type">optional object type</param>
		/// <param name="options">optional object options</param>
		[SupportByVersion("DAO", 3.6,12.0)]
		[BaseResult]
		NetOffice.DAOApi.Recordset _30_OpenRecordset(string name, object type, object options);

		/// <summary>
		/// SupportByVersion DAO 3.6, 12.0
		/// </summary>
		/// <param name="name">string name</param>
		[CustomMethod]
		[BaseResult]
		[SupportByVersion("DAO", 3.6,12.0)]
		NetOffice.DAOApi.Recordset _30_OpenRecordset(string name);

		/// <summary>
		/// SupportByVersion DAO 3.6, 12.0
		/// </summary>
		/// <param name="name">string name</param>
		/// <param name="type">optional object type</param>
		[CustomMethod]
		[BaseResult]
		[SupportByVersion("DAO", 3.6,12.0)]
		NetOffice.DAOApi.Recordset _30_OpenRecordset(string name, object type);

		/// <summary>
		/// SupportByVersion DAO 3.6, 12.0
		/// </summary>
		/// <param name="name">optional object name</param>
		/// <param name="type">optional object type</param>
		/// <param name="value">optional object value</param>
		/// <param name="dDL">optional object dDL</param>
		[SupportByVersion("DAO", 3.6,12.0)]
		NetOffice.DAOApi.Property CreateProperty(object name, object type, object value, object dDL);

		/// <summary>
		/// SupportByVersion DAO 3.6, 12.0
		/// </summary>
		[CustomMethod]
		[SupportByVersion("DAO", 3.6,12.0)]
		NetOffice.DAOApi.Property CreateProperty();

		/// <summary>
		/// SupportByVersion DAO 3.6, 12.0
		/// </summary>
		/// <param name="name">optional object name</param>
		[CustomMethod]
		[SupportByVersion("DAO", 3.6,12.0)]
		NetOffice.DAOApi.Property CreateProperty(object name);

		/// <summary>
		/// SupportByVersion DAO 3.6, 12.0
		/// </summary>
		/// <param name="name">optional object name</param>
		/// <param name="type">optional object type</param>
		[CustomMethod]
		[SupportByVersion("DAO", 3.6,12.0)]
		NetOffice.DAOApi.Property CreateProperty(object name, object type);

		/// <summary>
		/// SupportByVersion DAO 3.6, 12.0
		/// </summary>
		/// <param name="name">optional object name</param>
		/// <param name="type">optional object type</param>
		/// <param name="value">optional object value</param>
		[CustomMethod]
		[SupportByVersion("DAO", 3.6,12.0)]
		NetOffice.DAOApi.Property CreateProperty(object name, object type, object value);

		/// <summary>
		/// SupportByVersion DAO 3.6, 12.0
		/// </summary>
		/// <param name="name">optional object name</param>
		/// <param name="table">optional object table</param>
		/// <param name="foreignTable">optional object foreignTable</param>
		/// <param name="attributes">optional object attributes</param>
		[SupportByVersion("DAO", 3.6,12.0)]
		NetOffice.DAOApi.Relation CreateRelation(object name, object table, object foreignTable, object attributes);

		/// <summary>
		/// SupportByVersion DAO 3.6, 12.0
		/// </summary>
		[CustomMethod]
		[SupportByVersion("DAO", 3.6,12.0)]
		NetOffice.DAOApi.Relation CreateRelation();

		/// <summary>
		/// SupportByVersion DAO 3.6, 12.0
		/// </summary>
		/// <param name="name">optional object name</param>
		[CustomMethod]
		[SupportByVersion("DAO", 3.6,12.0)]
		NetOffice.DAOApi.Relation CreateRelation(object name);

		/// <summary>
		/// SupportByVersion DAO 3.6, 12.0
		/// </summary>
		/// <param name="name">optional object name</param>
		/// <param name="table">optional object table</param>
		[CustomMethod]
		[SupportByVersion("DAO", 3.6,12.0)]
		NetOffice.DAOApi.Relation CreateRelation(object name, object table);

		/// <summary>
		/// SupportByVersion DAO 3.6, 12.0
		/// </summary>
		/// <param name="name">optional object name</param>
		/// <param name="table">optional object table</param>
		/// <param name="foreignTable">optional object foreignTable</param>
		[CustomMethod]
		[SupportByVersion("DAO", 3.6,12.0)]
		NetOffice.DAOApi.Relation CreateRelation(object name, object table, object foreignTable);

		/// <summary>
		/// SupportByVersion DAO 3.6, 12.0
		/// </summary>
		/// <param name="name">optional object name</param>
		/// <param name="attributes">optional object attributes</param>
		/// <param name="sourceTableName">optional object sourceTableName</param>
		/// <param name="connect">optional object connect</param>
		[SupportByVersion("DAO", 3.6,12.0)]
		NetOffice.DAOApi.TableDef CreateTableDef(object name, object attributes, object sourceTableName, object connect);

		/// <summary>
		/// SupportByVersion DAO 3.6, 12.0
		/// </summary>
		[CustomMethod]
		[SupportByVersion("DAO", 3.6,12.0)]
		NetOffice.DAOApi.TableDef CreateTableDef();

		/// <summary>
		/// SupportByVersion DAO 3.6, 12.0
		/// </summary>
		/// <param name="name">optional object name</param>
		[CustomMethod]
		[SupportByVersion("DAO", 3.6,12.0)]
		NetOffice.DAOApi.TableDef CreateTableDef(object name);

		/// <summary>
		/// SupportByVersion DAO 3.6, 12.0
		/// </summary>
		/// <param name="name">optional object name</param>
		/// <param name="attributes">optional object attributes</param>
		[CustomMethod]
		[SupportByVersion("DAO", 3.6,12.0)]
		NetOffice.DAOApi.TableDef CreateTableDef(object name, object attributes);

		/// <summary>
		/// SupportByVersion DAO 3.6, 12.0
		/// </summary>
		/// <param name="name">optional object name</param>
		/// <param name="attributes">optional object attributes</param>
		/// <param name="sourceTableName">optional object sourceTableName</param>
		[CustomMethod]
		[SupportByVersion("DAO", 3.6,12.0)]
		NetOffice.DAOApi.TableDef CreateTableDef(object name, object attributes, object sourceTableName);

		/// <summary>
		/// SupportByVersion DAO 3.6, 12.0
		/// </summary>
		[SupportByVersion("DAO", 3.6,12.0)]
		void BeginTrans();

		/// <summary>
		/// SupportByVersion DAO 3.6, 12.0
		/// </summary>
		/// <param name="options">optional Int32 Options = 0</param>
		[SupportByVersion("DAO", 3.6,12.0)]
		void CommitTrans(object options);

		/// <summary>
		/// SupportByVersion DAO 3.6, 12.0
		/// </summary>
		[CustomMethod]
		[SupportByVersion("DAO", 3.6,12.0)]
		void CommitTrans();

		/// <summary>
		/// SupportByVersion DAO 3.6, 12.0
		/// </summary>
		[SupportByVersion("DAO", 3.6,12.0)]
		void Rollback();

		/// <summary>
		/// SupportByVersion DAO 3.6, 12.0
		/// </summary>
		/// <param name="name">string name</param>
		/// <param name="options">optional object options</param>
		/// <param name="inconsistent">optional object inconsistent</param>
		[SupportByVersion("DAO", 3.6,12.0)]
		[BaseResult]
		NetOffice.DAOApi.Recordset CreateDynaset(string name, object options, object inconsistent);

		/// <summary>
		/// SupportByVersion DAO 3.6, 12.0
		/// </summary>
		/// <param name="name">string name</param>
		[CustomMethod]
		[BaseResult]
		[SupportByVersion("DAO", 3.6,12.0)]
		NetOffice.DAOApi.Recordset CreateDynaset(string name);

		/// <summary>
		/// SupportByVersion DAO 3.6, 12.0
		/// </summary>
		/// <param name="name">string name</param>
		/// <param name="options">optional object options</param>
		[CustomMethod]
		[BaseResult]
		[SupportByVersion("DAO", 3.6,12.0)]
		NetOffice.DAOApi.Recordset CreateDynaset(string name, object options);

		/// <summary>
		/// SupportByVersion DAO 3.6, 12.0
		/// </summary>
		/// <param name="name">optional object name</param>
		/// <param name="sQLText">optional object sQLText</param>
		[SupportByVersion("DAO", 3.6,12.0)]
		NetOffice.DAOApi.QueryDef CreateQueryDef(object name, object sQLText);

		/// <summary>
		/// SupportByVersion DAO 3.6, 12.0
		/// </summary>
		[CustomMethod]
		[SupportByVersion("DAO", 3.6,12.0)]
		NetOffice.DAOApi.QueryDef CreateQueryDef();

		/// <summary>
		/// SupportByVersion DAO 3.6, 12.0
		/// </summary>
		/// <param name="name">optional object name</param>
		[CustomMethod]
		[SupportByVersion("DAO", 3.6,12.0)]
		NetOffice.DAOApi.QueryDef CreateQueryDef(object name);

		/// <summary>
		/// SupportByVersion DAO 3.6, 12.0
		/// </summary>
		/// <param name="source">string source</param>
		/// <param name="options">optional object options</param>
		[SupportByVersion("DAO", 3.6,12.0)]
		[BaseResult]
		NetOffice.DAOApi.Recordset CreateSnapshot(string source, object options);

		/// <summary>
		/// SupportByVersion DAO 3.6, 12.0
		/// </summary>
		/// <param name="source">string source</param>
		[CustomMethod]
		[BaseResult]
		[SupportByVersion("DAO", 3.6,12.0)]
		NetOffice.DAOApi.Recordset CreateSnapshot(string source);

		/// <summary>
		/// SupportByVersion DAO 3.6, 12.0
		/// </summary>
		/// <param name="name">string name</param>
		[SupportByVersion("DAO", 3.6,12.0)]
		void DeleteQueryDef(string name);

		/// <summary>
		/// SupportByVersion DAO 3.6, 12.0
		/// </summary>
		/// <param name="sQL">string sQL</param>
		[SupportByVersion("DAO", 3.6,12.0)]
		Int32 ExecuteSQL(string sQL);

		/// <summary>
		/// SupportByVersion DAO 3.6, 12.0
		/// </summary>
		/// <param name="name">string name</param>
		[SupportByVersion("DAO", 3.6,12.0)]
		[BaseResult]
		NetOffice.DAOApi.Recordset ListFields(string name);

		/// <summary>
		/// SupportByVersion DAO 3.6, 12.0
		/// </summary>
		[SupportByVersion("DAO", 3.6,12.0)]
		[BaseResult]
		NetOffice.DAOApi.Recordset ListTables();

		/// <summary>
		/// SupportByVersion DAO 3.6, 12.0
		/// </summary>
		/// <param name="name">string name</param>
		[SupportByVersion("DAO", 3.6,12.0)]
		NetOffice.DAOApi.QueryDef OpenQueryDef(string name);

		/// <summary>
		/// SupportByVersion DAO 3.6, 12.0
		/// </summary>
		/// <param name="name">string name</param>
		/// <param name="options">optional object options</param>
		[SupportByVersion("DAO", 3.6,12.0)]
		[BaseResult]
		NetOffice.DAOApi.Recordset OpenTable(string name, object options);

		/// <summary>
		/// SupportByVersion DAO 3.6, 12.0
		/// </summary>
		/// <param name="name">string name</param>
		[CustomMethod]
		[BaseResult]
		[SupportByVersion("DAO", 3.6,12.0)]
		NetOffice.DAOApi.Recordset OpenTable(string name);

		/// <summary>
		/// SupportByVersion DAO 3.6, 12.0
		/// </summary>
		/// <param name="dbPathName">string dbPathName</param>
		/// <param name="exchangeType">optional object exchangeType</param>
		[SupportByVersion("DAO", 3.6,12.0)]
		void Synchronize(string dbPathName, object exchangeType);

		/// <summary>
		/// SupportByVersion DAO 3.6, 12.0
		/// </summary>
		/// <param name="dbPathName">string dbPathName</param>
		[CustomMethod]
		[SupportByVersion("DAO", 3.6,12.0)]
		void Synchronize(string dbPathName);

		/// <summary>
		/// SupportByVersion DAO 3.6, 12.0
		/// </summary>
		/// <param name="pathName">string pathName</param>
		/// <param name="description">string description</param>
		/// <param name="options">optional object options</param>
		[SupportByVersion("DAO", 3.6,12.0)]
		void MakeReplica(string pathName, string description, object options);

		/// <summary>
		/// SupportByVersion DAO 3.6, 12.0
		/// </summary>
		/// <param name="pathName">string pathName</param>
		/// <param name="description">string description</param>
		[CustomMethod]
		[SupportByVersion("DAO", 3.6,12.0)]
		void MakeReplica(string pathName, string description);

		/// <summary>
		/// SupportByVersion DAO 3.6, 12.0
		/// </summary>
		/// <param name="bstrOld">string bstrOld</param>
		/// <param name="bstrNew">string bstrNew</param>
		[SupportByVersion("DAO", 3.6,12.0)]
		void NewPassword(string bstrOld, string bstrNew);

		/// <summary>
		/// SupportByVersion DAO 3.6, 12.0
		/// </summary>
		/// <param name="name">string name</param>
		/// <param name="type">optional object type</param>
		/// <param name="options">optional object options</param>
		/// <param name="lockEdit">optional object lockEdit</param>
		[SupportByVersion("DAO", 3.6,12.0)]
		[BaseResult]
		NetOffice.DAOApi.Recordset OpenRecordset(string name, object type, object options, object lockEdit);

		/// <summary>
		/// SupportByVersion DAO 3.6, 12.0
		/// </summary>
		/// <param name="name">string name</param>
		[CustomMethod]
		[BaseResult]
		[SupportByVersion("DAO", 3.6,12.0)]
		NetOffice.DAOApi.Recordset OpenRecordset(string name);

		/// <summary>
		/// SupportByVersion DAO 3.6, 12.0
		/// </summary>
		/// <param name="name">string name</param>
		/// <param name="type">optional object type</param>
		[CustomMethod]
		[BaseResult]
		[SupportByVersion("DAO", 3.6,12.0)]
		NetOffice.DAOApi.Recordset OpenRecordset(string name, object type);

		/// <summary>
		/// SupportByVersion DAO 3.6, 12.0
		/// </summary>
		/// <param name="name">string name</param>
		/// <param name="type">optional object type</param>
		/// <param name="options">optional object options</param>
		[CustomMethod]
		[BaseResult]
		[SupportByVersion("DAO", 3.6,12.0)]
		NetOffice.DAOApi.Recordset OpenRecordset(string name, object type, object options);

		/// <summary>
		/// SupportByVersion DAO 3.6, 12.0
		/// </summary>
		/// <param name="dbPathName">string dbPathName</param>
		[SupportByVersion("DAO", 3.6,12.0)]
		void PopulatePartial(string dbPathName);

		#endregion
	}
}
