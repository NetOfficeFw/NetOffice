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
	/// DispatchInterface IVBDataObjectFiles 
	/// SupportByVersion MSComctlLib, 6
	/// </summary>
	[SupportByVersion("MSComctlLib", 6)]
	[EntityType(EntityType.IsDispatchInterface), BaseType, Enumerator(Enumerator.Value, EnumeratorInvoke.Method, "MSComctlLib", 6), HasIndexProperty(IndexInvoke.Property, "Item")]
	[TypeId("2334D2B3-713E-11CF-8AE5-00AA00C00905")]
    [CoClassSource(typeof(NetOffice.MSComctlLibApi.DataObjectFiles))]
    public interface IVBDataObjectFiles : ICOMObject, IEnumerableProvider<string>
	{
		#region Properties

		/// <summary>
		/// SupportByVersion MSComctlLib 6
		/// Get
		/// </summary>
		/// <param name="lIndex">Int32 lIndex</param>
		[SupportByVersion("MSComctlLib", 6)]
		[NetRuntimeSystem.Runtime.CompilerServices.IndexerName("Item"), IndexProperty]
		string this[Int32 lIndex] { get; }

		/// <summary>
		/// SupportByVersion MSComctlLib 6
		/// Get
		/// </summary>
		[SupportByVersion("MSComctlLib", 6)]
		Int32 Count { get; }

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion MSComctlLib 6
		/// </summary>
		/// <param name="bstrFilename">string bstrFilename</param>
		/// <param name="vIndex">optional object vIndex</param>
		[SupportByVersion("MSComctlLib", 6)]
		void Add(string bstrFilename, object vIndex);

		/// <summary>
		/// SupportByVersion MSComctlLib 6
		/// </summary>
		/// <param name="bstrFilename">string bstrFilename</param>
		[CustomMethod]
		[SupportByVersion("MSComctlLib", 6)]
		void Add(string bstrFilename);

		/// <summary>
		/// SupportByVersion MSComctlLib 6
		/// </summary>
		[SupportByVersion("MSComctlLib", 6)]
		void Clear();

		/// <summary>
		/// SupportByVersion MSComctlLib 6
		/// </summary>
		/// <param name="vIndex">object vIndex</param>
		[SupportByVersion("MSComctlLib", 6)]
		void Remove(object vIndex);

        #endregion

        #region IEnumerable<string>

        /// <summary>
        /// SupportByVersion MSComctlLib, 6
        /// </summary>
        [SupportByVersion("MSComctlLib", 6)]
        new IEnumerator<string> GetEnumerator();

        #endregion
    }
}
