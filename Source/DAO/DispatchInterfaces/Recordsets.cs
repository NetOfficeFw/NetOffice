using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.DAOApi
{
	/// <summary>
	/// DispatchInterface Recordsets 
	/// SupportByVersion DAO, 3.6,12.0
	/// </summary>
	[SupportByVersion("DAO", 3.6,12.0)]
	[EntityType(EntityType.IsDispatchInterface), HasIndexProperty(IndexInvoke.Property, "Item")]
	[TypeId("00000033-0000-0010-8000-00AA006D2EA4")]
	public interface Recordsets : _Collection
	{
		#region Properties

		/// <summary>
		/// SupportByVersion DAO 3.6, 12.0
		/// Get
		/// </summary>
		/// <param name="item">object item</param>
		[SupportByVersion("DAO", 3.6,12.0)]
		[BaseResult]
		[NetRuntimeSystem.Runtime.CompilerServices.IndexerName("Item"), IndexProperty]
		NetOffice.DAOApi.Recordset this[object item] { get; }

		#endregion

	}
}
