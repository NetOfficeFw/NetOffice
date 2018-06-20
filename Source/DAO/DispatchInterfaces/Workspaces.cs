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
	/// DispatchInterface Workspaces 
	/// SupportByVersion DAO, 3.6,12.0
	/// </summary>
	[SupportByVersion("DAO", 3.6,12.0)]
	[EntityType(EntityType.IsDispatchInterface), Enumerator(Enumerator.Reference, EnumeratorInvoke.Method, "DAO", 3.6, 12.0), HasIndexProperty(IndexInvoke.Property, "Item")]
	[TypeId("0000003B-0000-0010-8000-00AA006D2EA4")]
	public interface Workspaces : _DynaCollection, IEnumerableProvider<NetOffice.DAOApi.Workspace>
	{
		#region Properties

		/// <summary>
		/// SupportByVersion DAO 3.6, 12.0
		/// Get
		/// </summary>
		[SupportByVersion("DAO", 3.6,12.0)]
        new Int16 Count { get; }

		/// <summary>
		/// SupportByVersion DAO 3.6, 12.0
		/// Get
		/// </summary>
		/// <param name="item">object item</param>
		[SupportByVersion("DAO", 3.6,12.0)]
		[NetRuntimeSystem.Runtime.CompilerServices.IndexerName("Item"), IndexProperty]
		NetOffice.DAOApi.Workspace this[object item] { get; }

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion DAO 3.6, 12.0
		/// </summary>
		[SupportByVersion("DAO", 3.6,12.0)]
        new void Refresh();

        #endregion

        #region IEnumerable<NetOffice.DAOApi.Workspace>

        /// <summary>
        /// SupportByVersion DAO, 3.6,12.0
        /// </summary>
        [SupportByVersion("DAO", 3.6, 12.0)]
        new IEnumerator<NetOffice.DAOApi.Workspace> GetEnumerator();

        #endregion
    }
}
