using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.ADODBApi
{
	/// <summary>
	/// DispatchInterface Properties 
	/// SupportByVersion ADODB, 2.1,2.5
	/// </summary>
	[SupportByVersion("ADODB", 2.1,2.5)]
	[EntityType(EntityType.IsDispatchInterface), HasIndexProperty(IndexInvoke.Property, "Item")]
	[TypeId("00000504-0000-0010-8000-00AA006D2EA4")]
	public interface Properties : _Collection
	{
		#region Properties

		/// <summary>
		/// SupportByVersion ADODB 2.1, 2.5
		/// Get
		/// </summary>
		/// <param name="index">object index</param>
		[SupportByVersion("ADODB", 2.1,2.5)]
		[NetRuntimeSystem.Runtime.CompilerServices.IndexerName("Item"), IndexProperty]
		NetOffice.ADODBApi.Property this[object index] { get; }

		#endregion

	}
}
