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
	/// DispatchInterface Users 
	/// SupportByVersion DAO, 3.6,12.0
	/// </summary>
	[SupportByVersion("DAO", 3.6,12.0)]
	[EntityType(EntityType.IsDispatchInterface), Enumerator(Enumerator.Reference, EnumeratorInvoke.Method, "DAO", 3.6, 12.0), HasIndexProperty(IndexInvoke.Property, "Item")]
	[TypeId("0000006B-0000-0010-8000-00AA006D2EA4")]
	public interface Users : _DynaCollection, IEnumerableProvider<NetOffice.DAOApi.User>
	{
		#region Properties

		/// <summary>
		/// SupportByVersion DAO 3.6, 12.0
		/// Get
		/// </summary>
		/// <param name="item">object item</param>
		[SupportByVersion("DAO", 3.6,12.0)]
		[NetRuntimeSystem.Runtime.CompilerServices.IndexerName("Item"), IndexProperty]
		NetOffice.DAOApi.User this[object item] { get; }

        #endregion

        #region IEnumerable<NetOffice.DAOApi.User> Member

        /// <summary>
        /// SupportByVersion DAO, 3.6,12.0
        /// </summary>
        [SupportByVersion("DAO", 3.6, 12.0)]
        new IEnumerator<NetOffice.DAOApi.User> GetEnumerator();
        
        #endregion
    }
}
