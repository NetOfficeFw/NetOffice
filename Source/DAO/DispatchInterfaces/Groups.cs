using System.Collections;
using System.Collections.Generic;
using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;
using NetOffice.CollectionsGeneric;

namespace NetOffice.DAOApi
{
	/// <summary>
	/// DispatchInterface Groups 
	/// SupportByVersion DAO, 3.6,12.0
	/// </summary>
	[SupportByVersion("DAO", 3.6,12.0)]
	[EntityType(EntityType.IsDispatchInterface), Enumerator(Enumerator.Reference, EnumeratorInvoke.Method, "DAO", 3.6, 12.0), HasIndexProperty(IndexInvoke.Property, "Item")]
	[TypeId("00000063-0000-0010-8000-00AA006D2EA4")]
	public interface Groups : _DynaCollection, IEnumerableProvider<NetOffice.DAOApi.Group>
	{
		#region Properties

		/// <summary>
		/// SupportByVersion DAO 3.6, 12.0
		/// Get
		/// </summary>
		/// <param name="item">object item</param>
		[SupportByVersion("DAO", 3.6,12.0)]
		[NetRuntimeSystem.Runtime.CompilerServices.IndexerName("Item"), IndexProperty]
		NetOffice.DAOApi.Group this[object item] { get; }

        #endregion

        #region IEnumerable<NetOffice.DAOApi.Group>

        /// <summary>
        /// SupportByVersion DAO, 3.6,12.0
        /// </summary>
        [SupportByVersion("DAO", 3.6, 12.0)]
        new IEnumerator<NetOffice.DAOApi.Group> GetEnumerator();

        #endregion
    }
}
