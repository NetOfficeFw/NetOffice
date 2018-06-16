using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.DAOApi
{
	/// <summary>
	/// DispatchInterface _DBEngine 
	/// SupportByVersion DAO, 3.6,12.0
	/// </summary>
	[SupportByVersion("DAO", 3.6,12.0)]
	[EntityType(EntityType.IsDispatchInterface), BaseType]
	[TypeId("00000021-0000-0010-8000-00AA006D2EA4")]
    [CoClassSource(typeof(NetOffice.DAOApi.DBEngine), typeof(NetOffice.DAOApi.PrivDBEngine))]
	public interface _DBEngine : _DAO
	{
		#region Properties

		/// <summary>
		/// SupportByVersion DAO 3.6, 12.0
		/// Get
		/// </summary>
		[SupportByVersion("DAO", 3.6,12.0)]
		string Version { get; }

		/// <summary>
		/// SupportByVersion DAO 3.6, 12.0
		/// Get/Set
		/// </summary>
		[SupportByVersion("DAO", 3.6,12.0)]
		string IniPath { get; set; }

		/// <summary>
		/// SupportByVersion DAO 3.6, 12.0
		/// Get/Set
		/// </summary>
		[SupportByVersion("DAO", 3.6,12.0)]
		string DefaultUser { get; set; }

		/// <summary>
		/// SupportByVersion DAO 3.6, 12.0
		/// Get/Set
		/// </summary>
		[SupportByVersion("DAO", 3.6,12.0)]
		string DefaultPassword { get; set; }

		/// <summary>
		/// SupportByVersion DAO 3.6, 12.0
		/// Get/Set
		/// </summary>
		[SupportByVersion("DAO", 3.6,12.0)]
		Int16 LoginTimeout { get; set; }

		/// <summary>
		/// SupportByVersion DAO 3.6, 12.0
		/// Get
		/// </summary>
		[SupportByVersion("DAO", 3.6,12.0)]
		NetOffice.DAOApi.Workspaces Workspaces { get; }

		/// <summary>
		/// SupportByVersion DAO 3.6, 12.0
		/// Get
		/// </summary>
		[SupportByVersion("DAO", 3.6,12.0)]
		NetOffice.DAOApi.Errors Errors { get; }

		/// <summary>
		/// SupportByVersion DAO 3.6, 12.0
		/// Get/Set
		/// </summary>
		[SupportByVersion("DAO", 3.6,12.0)]
		string SystemDB { get; set; }

		/// <summary>
		/// SupportByVersion DAO 3.6, 12.0
		/// Get/Set
		/// </summary>
		[SupportByVersion("DAO", 3.6,12.0)]
		Int32 DefaultType { get; set; }

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion DAO 3.6, 12.0
		/// </summary>
		/// <param name="action">optional object action</param>
		[SupportByVersion("DAO", 3.6,12.0)]
		void Idle(object action);

		/// <summary>
		/// SupportByVersion DAO 3.6, 12.0
		/// </summary>
		[CustomMethod]
		[SupportByVersion("DAO", 3.6,12.0)]
		void Idle();

		/// <summary>
		/// SupportByVersion DAO 3.6, 12.0
		/// </summary>
		/// <param name="srcName">string srcName</param>
		/// <param name="dstName">string dstName</param>
		/// <param name="dstLocale">optional object dstLocale</param>
		/// <param name="options">optional object options</param>
		/// <param name="srcLocale">optional object srcLocale</param>
		[SupportByVersion("DAO", 3.6,12.0)]
		void CompactDatabase(string srcName, string dstName, object dstLocale, object options, object srcLocale);

		/// <summary>
		/// SupportByVersion DAO 3.6, 12.0
		/// </summary>
		/// <param name="srcName">string srcName</param>
		/// <param name="dstName">string dstName</param>
		[CustomMethod]
		[SupportByVersion("DAO", 3.6,12.0)]
		void CompactDatabase(string srcName, string dstName);

		/// <summary>
		/// SupportByVersion DAO 3.6, 12.0
		/// </summary>
		/// <param name="srcName">string srcName</param>
		/// <param name="dstName">string dstName</param>
		/// <param name="dstLocale">optional object dstLocale</param>
		[CustomMethod]
		[SupportByVersion("DAO", 3.6,12.0)]
		void CompactDatabase(string srcName, string dstName, object dstLocale);

		/// <summary>
		/// SupportByVersion DAO 3.6, 12.0
		/// </summary>
		/// <param name="srcName">string srcName</param>
		/// <param name="dstName">string dstName</param>
		/// <param name="dstLocale">optional object dstLocale</param>
		/// <param name="options">optional object options</param>
		[CustomMethod]
		[SupportByVersion("DAO", 3.6,12.0)]
		void CompactDatabase(string srcName, string dstName, object dstLocale, object options);

		/// <summary>
		/// SupportByVersion DAO 3.6, 12.0
		/// </summary>
		/// <param name="name">string name</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("DAO", 3.6,12.0)]
		void RepairDatabase(string name);

		/// <summary>
		/// SupportByVersion DAO 3.6, 12.0
		/// </summary>
		/// <param name="dsn">string dsn</param>
		/// <param name="driver">string driver</param>
		/// <param name="silent">bool silent</param>
		/// <param name="attributes">string attributes</param>
		[SupportByVersion("DAO", 3.6,12.0)]
		void RegisterDatabase(string dsn, string driver, bool silent, string attributes);

		/// <summary>
		/// SupportByVersion DAO 3.6, 12.0
		/// </summary>
		/// <param name="name">string name</param>
		/// <param name="userName">string userName</param>
		/// <param name="password">string password</param>
		[SupportByVersion("DAO", 3.6,12.0)]
		NetOffice.DAOApi.Workspace _30_CreateWorkspace(string name, string userName, string password);

		/// <summary>
		/// SupportByVersion DAO 3.6, 12.0
		/// </summary>
		/// <param name="name">string name</param>
		/// <param name="options">optional object options</param>
		/// <param name="readOnly">optional object readOnly</param>
		/// <param name="connect">optional object connect</param>
		[SupportByVersion("DAO", 3.6,12.0)]
		NetOffice.DAOApi.Database OpenDatabase(string name, object options, object readOnly, object connect);

		/// <summary>
		/// SupportByVersion DAO 3.6, 12.0
		/// </summary>
		/// <param name="name">string name</param>
		[CustomMethod]
		[SupportByVersion("DAO", 3.6,12.0)]
		NetOffice.DAOApi.Database OpenDatabase(string name);

		/// <summary>
		/// SupportByVersion DAO 3.6, 12.0
		/// </summary>
		/// <param name="name">string name</param>
		/// <param name="options">optional object options</param>
		[CustomMethod]
		[SupportByVersion("DAO", 3.6,12.0)]
		NetOffice.DAOApi.Database OpenDatabase(string name, object options);

		/// <summary>
		/// SupportByVersion DAO 3.6, 12.0
		/// </summary>
		/// <param name="name">string name</param>
		/// <param name="options">optional object options</param>
		/// <param name="readOnly">optional object readOnly</param>
		[CustomMethod]
		[SupportByVersion("DAO", 3.6,12.0)]
		NetOffice.DAOApi.Database OpenDatabase(string name, object options, object readOnly);

		/// <summary>
		/// SupportByVersion DAO 3.6, 12.0
		/// </summary>
		/// <param name="name">string name</param>
		/// <param name="locale">string locale</param>
		/// <param name="option">optional object option</param>
		[SupportByVersion("DAO", 3.6,12.0)]
		NetOffice.DAOApi.Database CreateDatabase(string name, string locale, object option);

		/// <summary>
		/// SupportByVersion DAO 3.6, 12.0
		/// </summary>
		/// <param name="name">string name</param>
		/// <param name="locale">string locale</param>
		[CustomMethod]
		[SupportByVersion("DAO", 3.6,12.0)]
		NetOffice.DAOApi.Database CreateDatabase(string name, string locale);

		/// <summary>
		/// SupportByVersion DAO 3.6, 12.0
		/// </summary>
		[SupportByVersion("DAO", 3.6,12.0)]
		void FreeLocks();

		/// <summary>
		/// SupportByVersion DAO 3.6, 12.0
		/// </summary>
		[SupportByVersion("DAO", 3.6,12.0)]
		void BeginTrans();

		/// <summary>
		/// SupportByVersion DAO 3.6, 12.0
		/// </summary>
		/// <param name="option">optional Int32 Option = 0</param>
		[SupportByVersion("DAO", 3.6,12.0)]
		void CommitTrans(object option);

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
		/// <param name="password">string password</param>
		[SupportByVersion("DAO", 3.6,12.0)]
		void SetDefaultWorkspace(string name, string password);

		/// <summary>
		/// SupportByVersion DAO 3.6, 12.0
		/// </summary>
		/// <param name="option">Int16 option</param>
		/// <param name="value">object value</param>
		[SupportByVersion("DAO", 3.6,12.0)]
		void SetDataAccessOption(Int16 option, object value);

		/// <summary>
		/// SupportByVersion DAO 3.6, 12.0
		/// </summary>
		/// <param name="statNum">Int32 statNum</param>
		/// <param name="reset">optional object reset</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("DAO", 3.6,12.0)]
		Int32 ISAMStats(Int32 statNum, object reset);

		/// <summary>
		/// SupportByVersion DAO 3.6, 12.0
		/// </summary>
		/// <param name="statNum">Int32 statNum</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("DAO", 3.6,12.0)]
		Int32 ISAMStats(Int32 statNum);

		/// <summary>
		/// SupportByVersion DAO 3.6, 12.0
		/// </summary>
		/// <param name="name">string name</param>
		/// <param name="userName">string userName</param>
		/// <param name="password">string password</param>
		/// <param name="useType">optional object useType</param>
		[SupportByVersion("DAO", 3.6,12.0)]
		NetOffice.DAOApi.Workspace CreateWorkspace(string name, string userName, string password, object useType);

		/// <summary>
		/// SupportByVersion DAO 3.6, 12.0
		/// </summary>
		/// <param name="name">string name</param>
		/// <param name="userName">string userName</param>
		/// <param name="password">string password</param>
		[CustomMethod]
		[SupportByVersion("DAO", 3.6,12.0)]
		NetOffice.DAOApi.Workspace CreateWorkspace(string name, string userName, string password);

		/// <summary>
		/// SupportByVersion DAO 3.6, 12.0
		/// </summary>
		/// <param name="name">string name</param>
		/// <param name="options">optional object options</param>
		/// <param name="readOnly">optional object readOnly</param>
		/// <param name="connect">optional object connect</param>
		[SupportByVersion("DAO", 3.6,12.0)]
		NetOffice.DAOApi.Connection OpenConnection(string name, object options, object readOnly, object connect);

		/// <summary>
		/// SupportByVersion DAO 3.6, 12.0
		/// </summary>
		/// <param name="name">string name</param>
		[CustomMethod]
		[SupportByVersion("DAO", 3.6,12.0)]
		NetOffice.DAOApi.Connection OpenConnection(string name);

		/// <summary>
		/// SupportByVersion DAO 3.6, 12.0
		/// </summary>
		/// <param name="name">string name</param>
		/// <param name="options">optional object options</param>
		[CustomMethod]
		[SupportByVersion("DAO", 3.6,12.0)]
		NetOffice.DAOApi.Connection OpenConnection(string name, object options);

		/// <summary>
		/// SupportByVersion DAO 3.6, 12.0
		/// </summary>
		/// <param name="name">string name</param>
		/// <param name="options">optional object options</param>
		/// <param name="readOnly">optional object readOnly</param>
		[CustomMethod]
		[SupportByVersion("DAO", 3.6,12.0)]
		NetOffice.DAOApi.Connection OpenConnection(string name, object options, object readOnly);

		/// <summary>
		/// SupportByVersion DAO 3.6, 12.0
		/// </summary>
		/// <param name="option">Int32 option</param>
		/// <param name="value">object value</param>
		[SupportByVersion("DAO", 3.6,12.0)]
		void SetOption(Int32 option, object value);

		#endregion
	}
}
