using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.ADODBApi
{
	/// <summary>
	/// DispatchInterface Fields15 
	/// SupportByVersion ADODB, 2.1,2.5
	/// </summary>
	[SupportByVersion("ADODB", 2.1,2.5)]
	[EntityType(EntityType.IsDispatchInterface), BaseType, HasIndexProperty(IndexInvoke.Property, "Item")]
	[TypeId("00000506-0000-0010-8000-00AA006D2EA4")]
	public interface Fields15 : _Collection
	{
		#region Properties

		/// <summary>
		/// SupportByVersion ADODB 2.1, 2.5
		/// Get
		/// </summary>
		/// <param name="index">object index</param>
		[SupportByVersion("ADODB", 2.1,2.5)]
		[NetRuntimeSystem.Runtime.CompilerServices.IndexerName("Item"), IndexProperty]
		NetOffice.ADODBApi.Field this[object index] { get; }

		#endregion

	}
}
