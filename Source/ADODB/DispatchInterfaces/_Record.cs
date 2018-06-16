using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.ADODBApi
{
	/// <summary>
	/// DispatchInterface _Record 
	/// SupportByVersion ADODB, 2.5
	/// </summary>
	[SupportByVersion("ADODB", 2.5)]
	[EntityType(EntityType.IsDispatchInterface), BaseType]
	[TypeId("00000562-0000-0010-8000-00AA006D2EA4")]
    [CoClassSource(typeof(NetOffice.ADODBApi.Record))]
    public interface _Record : _ADO
	{
		#region Properties

		/// <summary>
		/// SupportByVersion ADODB 2.5
		/// Get/Set
		/// </summary>
		[SupportByVersion("ADODB", 2.5)]
		object ActiveConnection { get; set; }

		/// <summary>
		/// SupportByVersion ADODB 2.5
		/// Get
		/// </summary>
		[SupportByVersion("ADODB", 2.5)]
		NetOffice.ADODBApi.Enums.ObjectStateEnum State { get; }

		/// <summary>
		/// SupportByVersion ADODB 2.5
		/// Get/Set
		/// </summary>
		[SupportByVersion("ADODB", 2.5)]
		object Source { get; set; }

		/// <summary>
		/// SupportByVersion ADODB 2.5
		/// Get/Set
		/// </summary>
		[SupportByVersion("ADODB", 2.5)]
		NetOffice.ADODBApi.Enums.ConnectModeEnum Mode { get; set; }

		/// <summary>
		/// SupportByVersion ADODB 2.5
		/// Get
		/// </summary>
		[SupportByVersion("ADODB", 2.5)]
		string ParentURL { get; }

		/// <summary>
		/// SupportByVersion ADODB 2.5
		/// Get
		/// </summary>
		[SupportByVersion("ADODB", 2.5)]
		NetOffice.ADODBApi.Fields Fields { get; }

		/// <summary>
		/// SupportByVersion ADODB 2.5
		/// Get
		/// </summary>
		[SupportByVersion("ADODB", 2.5)]
		NetOffice.ADODBApi.Enums.RecordTypeEnum RecordType { get; }

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion ADODB 2.5
		/// </summary>
		/// <param name="source">optional string Source = </param>
		/// <param name="destination">optional string Destination = </param>
		/// <param name="userName">optional string userName</param>
		/// <param name="password">optional string password</param>
		/// <param name="options">optional NetOffice.ADODBApi.Enums.MoveRecordOptionsEnum Options = -1</param>
		/// <param name="async">optional bool async</param>
		[SupportByVersion("ADODB", 2.5)]
		string MoveRecord(object source, object destination, object userName, object password, object options, object async);

		/// <summary>
		/// SupportByVersion ADODB 2.5
		/// </summary>
		[CustomMethod]
		[SupportByVersion("ADODB", 2.5)]
		string MoveRecord();

		/// <summary>
		/// SupportByVersion ADODB 2.5
		/// </summary>
		/// <param name="source">optional string Source = </param>
		[CustomMethod]
		[SupportByVersion("ADODB", 2.5)]
		string MoveRecord(object source);

		/// <summary>
		/// SupportByVersion ADODB 2.5
		/// </summary>
		/// <param name="source">optional string Source = </param>
		/// <param name="destination">optional string Destination = </param>
		[CustomMethod]
		[SupportByVersion("ADODB", 2.5)]
		string MoveRecord(object source, object destination);

		/// <summary>
		/// SupportByVersion ADODB 2.5
		/// </summary>
		/// <param name="source">optional string Source = </param>
		/// <param name="destination">optional string Destination = </param>
		/// <param name="userName">optional string userName</param>
		[CustomMethod]
		[SupportByVersion("ADODB", 2.5)]
		string MoveRecord(object source, object destination, object userName);

		/// <summary>
		/// SupportByVersion ADODB 2.5
		/// </summary>
		/// <param name="source">optional string Source = </param>
		/// <param name="destination">optional string Destination = </param>
		/// <param name="userName">optional string userName</param>
		/// <param name="password">optional string password</param>
		[CustomMethod]
		[SupportByVersion("ADODB", 2.5)]
		string MoveRecord(object source, object destination, object userName, object password);

		/// <summary>
		/// SupportByVersion ADODB 2.5
		/// </summary>
		/// <param name="source">optional string Source = </param>
		/// <param name="destination">optional string Destination = </param>
		/// <param name="userName">optional string userName</param>
		/// <param name="password">optional string password</param>
		/// <param name="options">optional NetOffice.ADODBApi.Enums.MoveRecordOptionsEnum Options = -1</param>
		[CustomMethod]
		[SupportByVersion("ADODB", 2.5)]
		string MoveRecord(object source, object destination, object userName, object password, object options);

		/// <summary>
		/// SupportByVersion ADODB 2.5
		/// </summary>
		/// <param name="source">optional string Source = </param>
		/// <param name="destination">optional string Destination = </param>
		/// <param name="userName">optional string userName</param>
		/// <param name="password">optional string password</param>
		/// <param name="options">optional NetOffice.ADODBApi.Enums.CopyRecordOptionsEnum Options = -1</param>
		/// <param name="async">optional bool async</param>
		[SupportByVersion("ADODB", 2.5)]
		string CopyRecord(object source, object destination, object userName, object password, object options, object async);

		/// <summary>
		/// SupportByVersion ADODB 2.5
		/// </summary>
		[CustomMethod]
		[SupportByVersion("ADODB", 2.5)]
		string CopyRecord();

		/// <summary>
		/// SupportByVersion ADODB 2.5
		/// </summary>
		/// <param name="source">optional string Source = </param>
		[CustomMethod]
		[SupportByVersion("ADODB", 2.5)]
		string CopyRecord(object source);

		/// <summary>
		/// SupportByVersion ADODB 2.5
		/// </summary>
		/// <param name="source">optional string Source = </param>
		/// <param name="destination">optional string Destination = </param>
		[CustomMethod]
		[SupportByVersion("ADODB", 2.5)]
		string CopyRecord(object source, object destination);

		/// <summary>
		/// SupportByVersion ADODB 2.5
		/// </summary>
		/// <param name="source">optional string Source = </param>
		/// <param name="destination">optional string Destination = </param>
		/// <param name="userName">optional string userName</param>
		[CustomMethod]
		[SupportByVersion("ADODB", 2.5)]
		string CopyRecord(object source, object destination, object userName);

		/// <summary>
		/// SupportByVersion ADODB 2.5
		/// </summary>
		/// <param name="source">optional string Source = </param>
		/// <param name="destination">optional string Destination = </param>
		/// <param name="userName">optional string userName</param>
		/// <param name="password">optional string password</param>
		[CustomMethod]
		[SupportByVersion("ADODB", 2.5)]
		string CopyRecord(object source, object destination, object userName, object password);

		/// <summary>
		/// SupportByVersion ADODB 2.5
		/// </summary>
		/// <param name="source">optional string Source = </param>
		/// <param name="destination">optional string Destination = </param>
		/// <param name="userName">optional string userName</param>
		/// <param name="password">optional string password</param>
		/// <param name="options">optional NetOffice.ADODBApi.Enums.CopyRecordOptionsEnum Options = -1</param>
		[CustomMethod]
		[SupportByVersion("ADODB", 2.5)]
		string CopyRecord(object source, object destination, object userName, object password, object options);

		/// <summary>
		/// SupportByVersion ADODB 2.5
		/// </summary>
		/// <param name="source">optional string source</param>
		/// <param name="async">optional bool async</param>
		[SupportByVersion("ADODB", 2.5)]
		void DeleteRecord(object source, object async);

		/// <summary>
		/// SupportByVersion ADODB 2.5
		/// </summary>
		[CustomMethod]
		[SupportByVersion("ADODB", 2.5)]
		void DeleteRecord();

		/// <summary>
		/// SupportByVersion ADODB 2.5
		/// </summary>
		/// <param name="source">optional string source</param>
		[CustomMethod]
		[SupportByVersion("ADODB", 2.5)]
		void DeleteRecord(object source);

		/// <summary>
		/// SupportByVersion ADODB 2.5
		/// </summary>
		/// <param name="source">optional object source</param>
		/// <param name="activeConnection">optional object activeConnection</param>
		/// <param name="mode">optional NetOffice.ADODBApi.Enums.ConnectModeEnum mode</param>
		/// <param name="createOptions">optional NetOffice.ADODBApi.Enums.RecordCreateOptionsEnum CreateOptions = -1</param>
		/// <param name="options">optional NetOffice.ADODBApi.Enums.RecordOpenOptionsEnum Options = -1</param>
		/// <param name="userName">optional string userName</param>
		/// <param name="password">optional string password</param>
		[SupportByVersion("ADODB", 2.5)]
		void Open(object source, object activeConnection, object mode, object createOptions, object options, object userName, object password);

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
		/// <param name="mode">optional NetOffice.ADODBApi.Enums.ConnectModeEnum mode</param>
		[CustomMethod]
		[SupportByVersion("ADODB", 2.5)]
		void Open(object source, object activeConnection, object mode);

		/// <summary>
		/// SupportByVersion ADODB 2.5
		/// </summary>
		/// <param name="source">optional object source</param>
		/// <param name="activeConnection">optional object activeConnection</param>
		/// <param name="mode">optional NetOffice.ADODBApi.Enums.ConnectModeEnum mode</param>
		/// <param name="createOptions">optional NetOffice.ADODBApi.Enums.RecordCreateOptionsEnum CreateOptions = -1</param>
		[CustomMethod]
		[SupportByVersion("ADODB", 2.5)]
		void Open(object source, object activeConnection, object mode, object createOptions);

		/// <summary>
		/// SupportByVersion ADODB 2.5
		/// </summary>
		/// <param name="source">optional object source</param>
		/// <param name="activeConnection">optional object activeConnection</param>
		/// <param name="mode">optional NetOffice.ADODBApi.Enums.ConnectModeEnum mode</param>
		/// <param name="createOptions">optional NetOffice.ADODBApi.Enums.RecordCreateOptionsEnum CreateOptions = -1</param>
		/// <param name="options">optional NetOffice.ADODBApi.Enums.RecordOpenOptionsEnum Options = -1</param>
		[CustomMethod]
		[SupportByVersion("ADODB", 2.5)]
		void Open(object source, object activeConnection, object mode, object createOptions, object options);

		/// <summary>
		/// SupportByVersion ADODB 2.5
		/// </summary>
		/// <param name="source">optional object source</param>
		/// <param name="activeConnection">optional object activeConnection</param>
		/// <param name="mode">optional NetOffice.ADODBApi.Enums.ConnectModeEnum mode</param>
		/// <param name="createOptions">optional NetOffice.ADODBApi.Enums.RecordCreateOptionsEnum CreateOptions = -1</param>
		/// <param name="options">optional NetOffice.ADODBApi.Enums.RecordOpenOptionsEnum Options = -1</param>
		/// <param name="userName">optional string userName</param>
		[CustomMethod]
		[SupportByVersion("ADODB", 2.5)]
		void Open(object source, object activeConnection, object mode, object createOptions, object options, object userName);

		/// <summary>
		/// SupportByVersion ADODB 2.5
		/// </summary>
		[SupportByVersion("ADODB", 2.5)]
		void Close();

		/// <summary>
		/// SupportByVersion ADODB 2.5
		/// </summary>
		[SupportByVersion("ADODB", 2.5)]
		[BaseResult]
		NetOffice.ADODBApi._Recordset GetChildren();

		/// <summary>
		/// SupportByVersion ADODB 2.5
		/// </summary>
		[SupportByVersion("ADODB", 2.5)]
		void Cancel();

		#endregion
	}
}
