using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.ADODBApi
{
	/// <summary>
	/// DispatchInterface Connection15 
	/// SupportByVersion ADODB, 2.1,2.5
	/// </summary>
	[SupportByVersion("ADODB", 2.1,2.5)]
	[EntityType(EntityType.IsDispatchInterface), BaseType]
	[TypeId("00000515-0000-0010-8000-00AA006D2EA4")]
	public interface Connection15 : _ADO
	{
		#region Properties

		/// <summary>
		/// SupportByVersion ADODB 2.1, 2.5
		/// Get/Set
		/// </summary>
		[SupportByVersion("ADODB", 2.1,2.5)]
		string ConnectionString { get; set; }

		/// <summary>
		/// SupportByVersion ADODB 2.1, 2.5
		/// Get/Set
		/// </summary>
		[SupportByVersion("ADODB", 2.1,2.5)]
		Int32 CommandTimeout { get; set; }

		/// <summary>
		/// SupportByVersion ADODB 2.1, 2.5
		/// Get/Set
		/// </summary>
		[SupportByVersion("ADODB", 2.1,2.5)]
		Int32 ConnectionTimeout { get; set; }

		/// <summary>
		/// SupportByVersion ADODB 2.1, 2.5
		/// Get
		/// </summary>
		[SupportByVersion("ADODB", 2.1,2.5)]
		string Version { get; }

		/// <summary>
		/// SupportByVersion ADODB 2.1, 2.5
		/// Get
		/// </summary>
		[SupportByVersion("ADODB", 2.1,2.5)]
		NetOffice.ADODBApi.Errors Errors { get; }

		/// <summary>
		/// SupportByVersion ADODB 2.1, 2.5
		/// Get/Set
		/// </summary>
		[SupportByVersion("ADODB", 2.1,2.5)]
		string DefaultDatabase { get; set; }

		/// <summary>
		/// SupportByVersion ADODB 2.1, 2.5
		/// Get/Set
		/// </summary>
		[SupportByVersion("ADODB", 2.1,2.5)]
		NetOffice.ADODBApi.Enums.IsolationLevelEnum IsolationLevel { get; set; }

		/// <summary>
		/// SupportByVersion ADODB 2.1, 2.5
		/// Get/Set
		/// </summary>
		[SupportByVersion("ADODB", 2.1,2.5)]
		Int32 Attributes { get; set; }

		/// <summary>
		/// SupportByVersion ADODB 2.1, 2.5
		/// Get/Set
		/// </summary>
		[SupportByVersion("ADODB", 2.1,2.5)]
		NetOffice.ADODBApi.Enums.CursorLocationEnum CursorLocation { get; set; }

		/// <summary>
		/// SupportByVersion ADODB 2.1, 2.5
		/// Get/Set
		/// </summary>
		[SupportByVersion("ADODB", 2.1,2.5)]
		NetOffice.ADODBApi.Enums.ConnectModeEnum Mode { get; set; }

		/// <summary>
		/// SupportByVersion ADODB 2.1, 2.5
		/// Get/Set
		/// </summary>
		[SupportByVersion("ADODB", 2.1,2.5)]
		string Provider { get; set; }

		/// <summary>
		/// SupportByVersion ADODB 2.1, 2.5
		/// Get
		/// </summary>
		[SupportByVersion("ADODB", 2.1,2.5)]
		Int32 State { get; }

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion ADODB 2.1, 2.5
		/// </summary>
		[SupportByVersion("ADODB", 2.1,2.5)]
		void Close();

		/// <summary>
		/// SupportByVersion ADODB 2.1, 2.5
		/// </summary>
		/// <param name="commandText">string commandText</param>
		/// <param name="recordsAffected">object recordsAffected</param>
		/// <param name="options">optional Int32 Options = -1</param>
		[SupportByVersion("ADODB", 2.1,2.5)]
		[BaseResult]
		NetOffice.ADODBApi._Recordset Execute(string commandText, object recordsAffected, object options);

		/// <summary>
		/// SupportByVersion ADODB 2.1, 2.5
		/// </summary>
		/// <param name="commandText">string commandText</param>
		/// <param name="recordsAffected">object recordsAffected</param>
		[CustomMethod]
		[BaseResult]
		[SupportByVersion("ADODB", 2.1,2.5)]
		NetOffice.ADODBApi._Recordset Execute(string commandText, object recordsAffected);

		/// <summary>
		/// SupportByVersion ADODB 2.1, 2.5
		/// </summary>
		[SupportByVersion("ADODB", 2.1,2.5)]
		Int32 BeginTrans();

		/// <summary>
		/// SupportByVersion ADODB 2.1, 2.5
		/// </summary>
		[SupportByVersion("ADODB", 2.1,2.5)]
		void CommitTrans();

		/// <summary>
		/// SupportByVersion ADODB 2.1, 2.5
		/// </summary>
		[SupportByVersion("ADODB", 2.1,2.5)]
		void RollbackTrans();

		/// <summary>
		/// SupportByVersion ADODB 2.1, 2.5
		/// </summary>
		/// <param name="connectionString">optional string ConnectionString = </param>
		/// <param name="userID">optional string UserID = </param>
		/// <param name="password">optional string Password = </param>
		/// <param name="options">optional Int32 Options = -1</param>
		[SupportByVersion("ADODB", 2.1,2.5)]
		void Open(object connectionString, object userID, object password, object options);

		/// <summary>
		/// SupportByVersion ADODB 2.1, 2.5
		/// </summary>
		[CustomMethod]
		[SupportByVersion("ADODB", 2.1,2.5)]
		void Open();

		/// <summary>
		/// SupportByVersion ADODB 2.1, 2.5
		/// </summary>
		/// <param name="connectionString">optional string ConnectionString = </param>
		[CustomMethod]
		[SupportByVersion("ADODB", 2.1,2.5)]
		void Open(object connectionString);

		/// <summary>
		/// SupportByVersion ADODB 2.1, 2.5
		/// </summary>
		/// <param name="connectionString">optional string ConnectionString = </param>
		/// <param name="userID">optional string UserID = </param>
		[CustomMethod]
		[SupportByVersion("ADODB", 2.1,2.5)]
		void Open(object connectionString, object userID);

		/// <summary>
		/// SupportByVersion ADODB 2.1, 2.5
		/// </summary>
		/// <param name="connectionString">optional string ConnectionString = </param>
		/// <param name="userID">optional string UserID = </param>
		/// <param name="password">optional string Password = </param>
		[CustomMethod]
		[SupportByVersion("ADODB", 2.1,2.5)]
		void Open(object connectionString, object userID, object password);

		/// <summary>
		/// SupportByVersion ADODB 2.1, 2.5
		/// </summary>
		/// <param name="schema">NetOffice.ADODBApi.Enums.SchemaEnum schema</param>
		/// <param name="restrictions">optional object restrictions</param>
		/// <param name="schemaID">optional object schemaID</param>
		[SupportByVersion("ADODB", 2.1,2.5)]
		[BaseResult]
		NetOffice.ADODBApi._Recordset OpenSchema(NetOffice.ADODBApi.Enums.SchemaEnum schema, object restrictions, object schemaID);

		/// <summary>
		/// SupportByVersion ADODB 2.1, 2.5
		/// </summary>
		/// <param name="schema">NetOffice.ADODBApi.Enums.SchemaEnum schema</param>
		[CustomMethod]
		[BaseResult]
		[SupportByVersion("ADODB", 2.1,2.5)]
		NetOffice.ADODBApi._Recordset OpenSchema(NetOffice.ADODBApi.Enums.SchemaEnum schema);

		/// <summary>
		/// SupportByVersion ADODB 2.1, 2.5
		/// </summary>
		/// <param name="schema">NetOffice.ADODBApi.Enums.SchemaEnum schema</param>
		/// <param name="restrictions">optional object restrictions</param>
		[CustomMethod]
		[BaseResult]
		[SupportByVersion("ADODB", 2.1,2.5)]
		NetOffice.ADODBApi._Recordset OpenSchema(NetOffice.ADODBApi.Enums.SchemaEnum schema, object restrictions);

		#endregion
	}
}
