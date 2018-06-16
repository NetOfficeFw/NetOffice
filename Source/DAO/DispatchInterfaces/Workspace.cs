using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.DAOApi
{
	/// <summary>
	/// DispatchInterface Workspace 
	/// SupportByVersion DAO, 3.6,12.0
	/// </summary>
	[SupportByVersion("DAO", 3.6,12.0)]
	[EntityType(EntityType.IsDispatchInterface)]
	[TypeId("00000039-0000-0010-8000-00AA006D2EA4")]
	public interface Workspace : _DAO
	{
		#region Properties

		/// <summary>
		/// SupportByVersion DAO 3.6, 12.0
		/// Get/Set
		/// </summary>
		[SupportByVersion("DAO", 3.6,12.0)]
		string Name { get; set; }

		/// <summary>
		/// SupportByVersion DAO 3.6, 12.0
		/// Get
		/// </summary>
		[SupportByVersion("DAO", 3.6,12.0)]
		string UserName { get; }

		/// <summary>
		/// SupportByVersion DAO 3.6, 12.0
		/// Get/Set
		/// </summary>
		[SupportByVersion("DAO", 3.6,12.0)]
		string _30_UserName { get; set; }

		/// <summary>
		/// SupportByVersion DAO 3.6, 12.0
		/// Get/Set
		/// </summary>
		[SupportByVersion("DAO", 3.6,12.0)]
		string _30_Password { get; set; }

		/// <summary>
		/// SupportByVersion DAO 3.6, 12.0
		/// Get/Set
		/// </summary>
		[SupportByVersion("DAO", 3.6,12.0)]
		Int16 IsolateODBCTrans { get; set; }

		/// <summary>
		/// SupportByVersion DAO 3.6, 12.0
		/// Get
		/// </summary>
		[SupportByVersion("DAO", 3.6,12.0)]
		NetOffice.DAOApi.Databases Databases { get; }

		/// <summary>
		/// SupportByVersion DAO 3.6, 12.0
		/// Get
		/// </summary>
		[SupportByVersion("DAO", 3.6,12.0)]
		NetOffice.DAOApi.Users Users { get; }

		/// <summary>
		/// SupportByVersion DAO 3.6, 12.0
		/// Get
		/// </summary>
		[SupportByVersion("DAO", 3.6,12.0)]
		NetOffice.DAOApi.Groups Groups { get; }

		/// <summary>
		/// SupportByVersion DAO 3.6, 12.0
		/// Get/Set
		/// </summary>
		[SupportByVersion("DAO", 3.6,12.0)]
		Int32 LoginTimeout { get; set; }

		/// <summary>
		/// SupportByVersion DAO 3.6, 12.0
		/// Get/Set
		/// </summary>
		[SupportByVersion("DAO", 3.6,12.0)]
		Int32 DefaultCursorDriver { get; set; }

		/// <summary>
		/// SupportByVersion DAO 3.6, 12.0
		/// Get
		/// </summary>
		[SupportByVersion("DAO", 3.6,12.0)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		Int32 hEnv { get; }

		/// <summary>
		/// SupportByVersion DAO 3.6, 12.0
		/// Get
		/// </summary>
		[SupportByVersion("DAO", 3.6,12.0)]
		Int32 Type { get; }

		/// <summary>
		/// SupportByVersion DAO 3.6, 12.0
		/// Get
		/// </summary>
		[SupportByVersion("DAO", 3.6,12.0)]
		NetOffice.DAOApi.Connections Connections { get; }

		#endregion

		#region Methods

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
		void Close();

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
		/// <param name="connect">string connect</param>
		/// <param name="option">optional object option</param>
		[SupportByVersion("DAO", 3.6,12.0)]
		NetOffice.DAOApi.Database CreateDatabase(string name, string connect, object option);

		/// <summary>
		/// SupportByVersion DAO 3.6, 12.0
		/// </summary>
		/// <param name="name">string name</param>
		/// <param name="connect">string connect</param>
		[CustomMethod]
		[SupportByVersion("DAO", 3.6,12.0)]
		NetOffice.DAOApi.Database CreateDatabase(string name, string connect);

		/// <summary>
		/// SupportByVersion DAO 3.6, 12.0
		/// </summary>
		/// <param name="name">optional object name</param>
		/// <param name="pID">optional object pID</param>
		/// <param name="password">optional object password</param>
		[SupportByVersion("DAO", 3.6,12.0)]
		NetOffice.DAOApi.User CreateUser(object name, object pID, object password);

		/// <summary>
		/// SupportByVersion DAO 3.6, 12.0
		/// </summary>
		[CustomMethod]
		[SupportByVersion("DAO", 3.6,12.0)]
		NetOffice.DAOApi.User CreateUser();

		/// <summary>
		/// SupportByVersion DAO 3.6, 12.0
		/// </summary>
		/// <param name="name">optional object name</param>
		[CustomMethod]
		[SupportByVersion("DAO", 3.6,12.0)]
		NetOffice.DAOApi.User CreateUser(object name);

		/// <summary>
		/// SupportByVersion DAO 3.6, 12.0
		/// </summary>
		/// <param name="name">optional object name</param>
		/// <param name="pID">optional object pID</param>
		[CustomMethod]
		[SupportByVersion("DAO", 3.6,12.0)]
		NetOffice.DAOApi.User CreateUser(object name, object pID);

		/// <summary>
		/// SupportByVersion DAO 3.6, 12.0
		/// </summary>
		/// <param name="name">optional object name</param>
		/// <param name="pID">optional object pID</param>
		[SupportByVersion("DAO", 3.6,12.0)]
		NetOffice.DAOApi.Group CreateGroup(object name, object pID);

		/// <summary>
		/// SupportByVersion DAO 3.6, 12.0
		/// </summary>
		[CustomMethod]
		[SupportByVersion("DAO", 3.6,12.0)]
		NetOffice.DAOApi.Group CreateGroup();

		/// <summary>
		/// SupportByVersion DAO 3.6, 12.0
		/// </summary>
		/// <param name="name">optional object name</param>
		[CustomMethod]
		[SupportByVersion("DAO", 3.6,12.0)]
		NetOffice.DAOApi.Group CreateGroup(object name);

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

		#endregion
	}
}
