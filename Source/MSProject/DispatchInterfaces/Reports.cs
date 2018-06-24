using System.Collections;
using System.Collections.Generic;
using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;
using NetOffice.CollectionsGeneric;

namespace NetOffice.MSProjectApi
{
	/// <summary>
	/// DispatchInterface Reports 
	/// SupportByVersion MSProject, 11
	/// </summary>
	[SupportByVersion("MSProject", 11)]
	[EntityType(EntityType.IsDispatchInterface), Enumerator(Enumerator.Reference, EnumeratorInvoke.Property, "MSProject", 11), HasIndexProperty(IndexInvoke.Property, "Item")]
	[TypeId("FF59CFBA-CB6F-4B92-A7D2-97D1CAB6EBFF")]
	public interface Reports : ICOMObject, IEnumerableProvider<NetOffice.MSProjectApi.Report>
	{
		#region Properties

		/// <summary>
		/// SupportByVersion MSProject 11
		/// Get
		/// </summary>
		/// <param name="index">object index</param>
		[SupportByVersion("MSProject", 11)]
		[NetRuntimeSystem.Runtime.CompilerServices.IndexerName("Item"), IndexProperty]
		NetOffice.MSProjectApi.Report this[object index] { get; }

		/// <summary>
		/// SupportByVersion MSProject 11
		/// Get
		/// </summary>
		[SupportByVersion("MSProject", 11)]
		Int32 Count { get; }

		/// <summary>
		/// SupportByVersion MSProject 11
		/// Get
		/// </summary>
		[SupportByVersion("MSProject", 11)]
		NetOffice.MSProjectApi.Project Parent { get; }

		/// <summary>
		/// SupportByVersion MSProject 11
		/// Get
		/// </summary>
		[SupportByVersion("MSProject", 11)]
		NetOffice.MSProjectApi.Application Application { get; }

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion MSProject 11
		/// </summary>
		/// <param name="source">object source</param>
		/// <param name="newName">string newName</param>
		[SupportByVersion("MSProject", 11)]
		NetOffice.MSProjectApi.Report Copy(object source, string newName);

		/// <summary>
		/// SupportByVersion MSProject 11
		/// </summary>
		/// <param name="name">string name</param>
		[SupportByVersion("MSProject", 11)]
		NetOffice.MSProjectApi.Report Add(string name);

		/// <summary>
		/// SupportByVersion MSProject 11
		/// </summary>
		/// <param name="name">string name</param>
		[SupportByVersion("MSProject", 11)]
		bool IsPresent(string name);

        #endregion

        #region IEnumerable<NetOffice.MSProjectApi.Report>

        /// <summary>
        /// SupportByVersion MSProject, 11
        /// </summary>
        [SupportByVersion("MSProject", 11)]
        new IEnumerator<NetOffice.MSProjectApi.Report> GetEnumerator();

        #endregion
    }
}
