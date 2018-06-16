using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.ADODBApi
{
	/// <summary>
	/// DispatchInterface ConnectionEvents_Deprecated 
	/// SupportByVersion ADODB, 2.5
	/// </summary>
	[SupportByVersion("ADODB", 2.5)]
	[EntityType(EntityType.IsDispatchInterface)]
	[TypeId("00000400-0000-0010-8000-00AA006D2EA4")]
	public interface ConnectionEvents_Deprecated : ICOMObject
	{
		#region Methods

		/// <summary>
		/// SupportByVersion ADODB 2.5
		/// </summary>
		/// <param name="pError">NetOffice.ADODBApi.Error pError</param>
		/// <param name="adStatus">NetOffice.ADODBApi.Enums.EventStatusEnum adStatus</param>
		/// <param name="pConnection">NetOffice.ADODBApi._Connection_Deprecated pConnection</param>
		[SupportByVersion("ADODB", 2.5)]
		Int32 InfoMessage(NetOffice.ADODBApi.Error pError, NetOffice.ADODBApi.Enums.EventStatusEnum adStatus, NetOffice.ADODBApi._Connection_Deprecated pConnection);

		/// <summary>
		/// SupportByVersion ADODB 2.5
		/// </summary>
		/// <param name="transactionLevel">Int32 transactionLevel</param>
		/// <param name="pError">NetOffice.ADODBApi.Error pError</param>
		/// <param name="adStatus">NetOffice.ADODBApi.Enums.EventStatusEnum adStatus</param>
		/// <param name="pConnection">NetOffice.ADODBApi._Connection_Deprecated pConnection</param>
		[SupportByVersion("ADODB", 2.5)]
		Int32 BeginTransComplete(Int32 transactionLevel, NetOffice.ADODBApi.Error pError, NetOffice.ADODBApi.Enums.EventStatusEnum adStatus, NetOffice.ADODBApi._Connection_Deprecated pConnection);

		/// <summary>
		/// SupportByVersion ADODB 2.5
		/// </summary>
		/// <param name="pError">NetOffice.ADODBApi.Error pError</param>
		/// <param name="adStatus">NetOffice.ADODBApi.Enums.EventStatusEnum adStatus</param>
		/// <param name="pConnection">NetOffice.ADODBApi._Connection_Deprecated pConnection</param>
		[SupportByVersion("ADODB", 2.5)]
		Int32 CommitTransComplete(NetOffice.ADODBApi.Error pError, NetOffice.ADODBApi.Enums.EventStatusEnum adStatus, NetOffice.ADODBApi._Connection_Deprecated pConnection);

		/// <summary>
		/// SupportByVersion ADODB 2.5
		/// </summary>
		/// <param name="pError">NetOffice.ADODBApi.Error pError</param>
		/// <param name="adStatus">NetOffice.ADODBApi.Enums.EventStatusEnum adStatus</param>
		/// <param name="pConnection">NetOffice.ADODBApi._Connection_Deprecated pConnection</param>
		[SupportByVersion("ADODB", 2.5)]
		Int32 RollbackTransComplete(NetOffice.ADODBApi.Error pError, NetOffice.ADODBApi.Enums.EventStatusEnum adStatus, NetOffice.ADODBApi._Connection_Deprecated pConnection);

		/// <summary>
		/// SupportByVersion ADODB 2.5
		/// </summary>
		/// <param name="source">string source</param>
		/// <param name="cursorType">NetOffice.ADODBApi.Enums.CursorTypeEnum cursorType</param>
		/// <param name="lockType">NetOffice.ADODBApi.Enums.LockTypeEnum lockType</param>
		/// <param name="options">Int32 options</param>
		/// <param name="adStatus">NetOffice.ADODBApi.Enums.EventStatusEnum adStatus</param>
		/// <param name="pCommand">NetOffice.ADODBApi._Command_Deprecated pCommand</param>
		/// <param name="pRecordset">NetOffice.ADODBApi._Recordset_Deprecated pRecordset</param>
		/// <param name="pConnection">NetOffice.ADODBApi._Connection_Deprecated pConnection</param>
		[SupportByVersion("ADODB", 2.5)]
		Int32 WillExecute(string source, NetOffice.ADODBApi.Enums.CursorTypeEnum cursorType, NetOffice.ADODBApi.Enums.LockTypeEnum lockType, Int32 options, NetOffice.ADODBApi.Enums.EventStatusEnum adStatus, NetOffice.ADODBApi._Command_Deprecated pCommand, NetOffice.ADODBApi._Recordset_Deprecated pRecordset, NetOffice.ADODBApi._Connection_Deprecated pConnection);

		/// <summary>
		/// SupportByVersion ADODB 2.5
		/// </summary>
		/// <param name="recordsAffected">Int32 recordsAffected</param>
		/// <param name="pError">NetOffice.ADODBApi.Error pError</param>
		/// <param name="adStatus">NetOffice.ADODBApi.Enums.EventStatusEnum adStatus</param>
		/// <param name="pCommand">NetOffice.ADODBApi._Command_Deprecated pCommand</param>
		/// <param name="pRecordset">NetOffice.ADODBApi._Recordset_Deprecated pRecordset</param>
		/// <param name="pConnection">NetOffice.ADODBApi._Connection_Deprecated pConnection</param>
		[SupportByVersion("ADODB", 2.5)]
		Int32 ExecuteComplete(Int32 recordsAffected, NetOffice.ADODBApi.Error pError, NetOffice.ADODBApi.Enums.EventStatusEnum adStatus, NetOffice.ADODBApi._Command_Deprecated pCommand, NetOffice.ADODBApi._Recordset_Deprecated pRecordset, NetOffice.ADODBApi._Connection_Deprecated pConnection);

		/// <summary>
		/// SupportByVersion ADODB 2.5
		/// </summary>
		/// <param name="connectionString">string connectionString</param>
		/// <param name="userID">string userID</param>
		/// <param name="password">string password</param>
		/// <param name="options">Int32 options</param>
		/// <param name="adStatus">NetOffice.ADODBApi.Enums.EventStatusEnum adStatus</param>
		/// <param name="pConnection">NetOffice.ADODBApi._Connection_Deprecated pConnection</param>
		[SupportByVersion("ADODB", 2.5)]
		Int32 WillConnect(string connectionString, string userID, string password, Int32 options, NetOffice.ADODBApi.Enums.EventStatusEnum adStatus, NetOffice.ADODBApi._Connection_Deprecated pConnection);

		/// <summary>
		/// SupportByVersion ADODB 2.5
		/// </summary>
		/// <param name="pError">NetOffice.ADODBApi.Error pError</param>
		/// <param name="adStatus">NetOffice.ADODBApi.Enums.EventStatusEnum adStatus</param>
		/// <param name="pConnection">NetOffice.ADODBApi._Connection_Deprecated pConnection</param>
		[SupportByVersion("ADODB", 2.5)]
		Int32 ConnectComplete(NetOffice.ADODBApi.Error pError, NetOffice.ADODBApi.Enums.EventStatusEnum adStatus, NetOffice.ADODBApi._Connection_Deprecated pConnection);

		/// <summary>
		/// SupportByVersion ADODB 2.5
		/// </summary>
		/// <param name="adStatus">NetOffice.ADODBApi.Enums.EventStatusEnum adStatus</param>
		/// <param name="pConnection">NetOffice.ADODBApi._Connection_Deprecated pConnection</param>
		[SupportByVersion("ADODB", 2.5)]
		Int32 Disconnect(NetOffice.ADODBApi.Enums.EventStatusEnum adStatus, NetOffice.ADODBApi._Connection_Deprecated pConnection);

		#endregion
	}
}
