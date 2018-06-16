using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.DAOApi
{
	/// <summary>
	/// DispatchInterface _Group 
	/// SupportByVersion DAO, 3.6,12.0
	/// </summary>
	[SupportByVersion("DAO", 3.6,12.0)]
	[EntityType(EntityType.IsDispatchInterface), BaseType]
	[TypeId("00000061-0000-0010-8000-00AA006D2EA4")]
    [CoClassSource(typeof(NetOffice.DAOApi.Group))]
    public interface _Group : _DAO
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
		/// Get/Set
		/// </summary>
		[SupportByVersion("DAO", 3.6,12.0)]
		string PID { get; set; }

		/// <summary>
		/// SupportByVersion DAO 3.6, 12.0
		/// Get
		/// </summary>
		[SupportByVersion("DAO", 3.6,12.0)]
		NetOffice.DAOApi.Users Users { get; }

		#endregion

		#region Methods

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

		#endregion
	}
}
