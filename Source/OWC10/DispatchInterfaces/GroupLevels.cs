using System.Collections;
using System.Collections.Generic;
using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;
using NetOffice.CollectionsGeneric;

namespace NetOffice.OWC10Api
{
	/// <summary>
	/// DispatchInterface GroupLevels 
	/// SupportByVersion OWC10, 1
	/// </summary>
	[SupportByVersion("OWC10", 1)]
	[EntityType(EntityType.IsDispatchInterface), Enumerator(Enumerator.Reference, EnumeratorInvoke.Property, "OWC10", 1), HasIndexProperty(IndexInvoke.Property, "Item")]
	[TypeId("F5B39B00-1480-11D3-8549-00C04FAC67D7")]
	public interface GroupLevels : ICOMObject, IEnumerableProvider<NetOffice.OWC10Api.GroupLevel>
	{
		#region Properties

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get
		/// </summary>
		/// <param name="index">object index</param>
		[SupportByVersion("OWC10", 1)]
		[NetRuntimeSystem.Runtime.CompilerServices.IndexerName("Item"), IndexProperty]
		NetOffice.OWC10Api.GroupLevel this[object index] { get; }

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		Int32 Count { get; }

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		/// <param name="recordSource">string recordSource</param>
		/// <param name="failIfThere">optional bool FailIfThere = false</param>
		[SupportByVersion("OWC10", 1)]
		NetOffice.OWC10Api.GroupLevel Add(string recordSource, object failIfThere);

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		/// <param name="recordSource">string recordSource</param>
		[CustomMethod]
		[SupportByVersion("OWC10", 1)]
		NetOffice.OWC10Api.GroupLevel Add(string recordSource);

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		/// <param name="index">object index</param>
		[SupportByVersion("OWC10", 1)]
		void Delete(object index);

        #endregion

        #region IEnumerable<NetOffice.OWC10Api.GroupLevel>

        /// <summary>
        /// SupportByVersion OWC10, 1
        /// </summary>
        [SupportByVersion("OWC10", 1)]
        new IEnumerator<NetOffice.OWC10Api.GroupLevel> GetEnumerator();

        #endregion
    }
}
