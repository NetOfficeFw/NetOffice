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
	/// DispatchInterface PayRates 
	/// SupportByVersion MSProject, 11,12,14
	/// </summary>
	/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff920650(v=office.14).aspx </remarks>
	[SupportByVersion("MSProject", 11,12,14)]
	[EntityType(EntityType.IsDispatchInterface), Enumerator(Enumerator.Reference, EnumeratorInvoke.Property, "MSProject", 11, 12, 14), HasIndexProperty(IndexInvoke.Property, "Item")]
	[TypeId("000C0C57-0000-0000-C000-000000000046")]
	public interface PayRates : ICOMObject, IEnumerableProvider<NetOffice.MSProjectApi.PayRate>
	{
		#region Properties

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get
		/// </summary>
		/// <param name="index">object index</param>
		[SupportByVersion("MSProject", 11,12,14)]
		[NetRuntimeSystem.Runtime.CompilerServices.IndexerName("Item"), IndexProperty]
		NetOffice.MSProjectApi.PayRate this[object index] { get; }

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		Int32 Count { get; }

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		NetOffice.MSProjectApi.Resource Parent { get; }

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		NetOffice.MSProjectApi.Application Application { get; }

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// </summary>
		/// <param name="effectiveDate">object effectiveDate</param>
		/// <param name="stdRate">optional object stdRate</param>
		/// <param name="ovtRate">optional object ovtRate</param>
		/// <param name="costPerUse">optional object costPerUse</param>
		[SupportByVersion("MSProject", 11,12,14)]
		NetOffice.MSProjectApi.PayRate Add(object effectiveDate, object stdRate, object ovtRate, object costPerUse);

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// </summary>
		/// <param name="effectiveDate">object effectiveDate</param>
		[CustomMethod]
		[SupportByVersion("MSProject", 11,12,14)]
		NetOffice.MSProjectApi.PayRate Add(object effectiveDate);

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// </summary>
		/// <param name="effectiveDate">object effectiveDate</param>
		/// <param name="stdRate">optional object stdRate</param>
		[CustomMethod]
		[SupportByVersion("MSProject", 11,12,14)]
		NetOffice.MSProjectApi.PayRate Add(object effectiveDate, object stdRate);

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// </summary>
		/// <param name="effectiveDate">object effectiveDate</param>
		/// <param name="stdRate">optional object stdRate</param>
		/// <param name="ovtRate">optional object ovtRate</param>
		[CustomMethod]
		[SupportByVersion("MSProject", 11,12,14)]
		NetOffice.MSProjectApi.PayRate Add(object effectiveDate, object stdRate, object ovtRate);

        #endregion

        #region IEnumerable<NetOffice.MSProjectApi.PayRate>

        /// <summary>
        /// SupportByVersion MSProject, 11,12,14
        /// </summary>
        [SupportByVersion("MSProject", 11, 12, 14)]
        new IEnumerator<NetOffice.MSProjectApi.PayRate> GetEnumerator();

        #endregion
    }
}
