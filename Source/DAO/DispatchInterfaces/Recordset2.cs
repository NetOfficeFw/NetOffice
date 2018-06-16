using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.DAOApi
{
	/// <summary>
	/// DispatchInterface Recordset2 
	/// SupportByVersion DAO, 12.0
	/// </summary>
	[SupportByVersion("DAO", 12.0)]
	[EntityType(EntityType.IsDispatchInterface)]
	[TypeId("00000035-0000-0010-8000-00AA006D2EA4")]
	public interface Recordset2 : Recordset
	{
		#region Properties

		/// <summary>
		/// SupportByVersion DAO 12.0
		/// Get
		/// </summary>
		[SupportByVersion("DAO", 12.0)]
        new NetOffice.DAOApi.Properties Properties { get; }

		/// <summary>
		/// SupportByVersion DAO 12.0
		/// Get
		/// </summary>
		[SupportByVersion("DAO", 12.0)]
		[BaseResult]
		NetOffice.DAOApi.Recordset ParentRecordset { get; }

		#endregion

	}
}
