using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.DAOApi
{
	/// <summary>
	/// DispatchInterface ComplexType 
	/// SupportByVersion DAO, 12.0
	/// </summary>
	[SupportByVersion("DAO", 12.0)]
	[EntityType(EntityType.IsDispatchInterface)]
	[TypeId("0000009D-0000-0010-8000-00AA006D2EA4")]
	public interface ComplexType : ICOMObject
	{
		#region Properties

		/// <summary>
		/// SupportByVersion DAO 12.0
		/// Get
		/// </summary>
		[SupportByVersion("DAO", 12.0)]
		NetOffice.DAOApi.Fields Fields { get; }

		#endregion

	}
}
