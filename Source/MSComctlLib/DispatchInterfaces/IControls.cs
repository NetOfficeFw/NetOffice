using System.Collections;
using System.Collections.Generic;
using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;
using NetOffice.CollectionsGeneric;

namespace NetOffice.MSComctlLibApi
{
	/// <summary>
	/// DispatchInterface IControls 
	/// SupportByVersion MSComctlLib, 6
	/// </summary>
	[SupportByVersion("MSComctlLib", 6)]
	[EntityType(EntityType.IsDispatchInterface), BaseType, Enumerator(Enumerator.Reference, EnumeratorInvoke.Method, "MSComctlLib", 6), HasIndexProperty(IndexInvoke.Property, "Item")]
	[TypeId("C8A3DC00-8593-11D1-B16A-00C0F0283628")]
    [CoClassSource(typeof(NetOffice.MSComctlLibApi.Controls))]
    public interface IControls : ICOMObject, IEnumerableProvider<object>
	{
		#region Properties

		/// <summary>
		/// SupportByVersion MSComctlLib 6
		/// Get
		/// </summary>
		[SupportByVersion("MSComctlLib", 6)]
		Int32 Count { get; }

		/// <summary>
		/// SupportByVersion MSComctlLib 6
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		/// <param name="index">Int32 index</param>
		[SupportByVersion("MSComctlLib", 6), ProxyResult]
		[NetRuntimeSystem.Runtime.CompilerServices.IndexerName("Item"), IndexProperty]
		object this[Int32 index] { get; }

        #endregion

        #region IEnumerable<object>

        /// <summary>
        /// SupportByVersion MSComctlLib, 6
        /// </summary>
        [SupportByVersion("MSComctlLib", 6)]
        new IEnumerator<object> GetEnumerator();

        #endregion
    }
}
