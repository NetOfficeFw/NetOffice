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
	/// DispatchInterface RecordsetDefs 
	/// SupportByVersion OWC10, 1
	/// </summary>
	[SupportByVersion("OWC10", 1)]
	[EntityType(EntityType.IsDispatchInterface), Enumerator(Enumerator.Reference, EnumeratorInvoke.Property, "OWC10", 1), HasIndexProperty(IndexInvoke.Property, "Item")]
	[TypeId("F5B39AA2-1480-11D3-8549-00C04FAC67D7")]
	public interface RecordsetDefs : ICOMObject, IEnumerableProvider<NetOffice.OWC10Api.RecordsetDef>
	{
		#region Properties

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		Int32 Count { get; }

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get
		/// </summary>
		/// <param name="index">object index</param>
		[SupportByVersion("OWC10", 1)]
		[NetRuntimeSystem.Runtime.CompilerServices.IndexerName("Item"), IndexProperty]
		NetOffice.OWC10Api.RecordsetDef this[object index] { get; }

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		/// <param name="schemaRowsource">object schemaRowsource</param>
		/// <param name="name">optional object name</param>
		[SupportByVersion("OWC10", 1)]
		NetOffice.OWC10Api.RecordsetDef Add(object schemaRowsource, object name);

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		/// <param name="schemaRowsource">object schemaRowsource</param>
		[CustomMethod]
		[SupportByVersion("OWC10", 1)]
		NetOffice.OWC10Api.RecordsetDef Add(object schemaRowsource);

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		/// <param name="source">string source</param>
		/// <param name="rowsourceType">optional object rowsourceType</param>
		/// <param name="name">optional object name</param>
		[SupportByVersion("OWC10", 1)]
		NetOffice.OWC10Api.RecordsetDef AddNew(string source, object rowsourceType, object name);

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		/// <param name="source">string source</param>
		[CustomMethod]
		[SupportByVersion("OWC10", 1)]
		NetOffice.OWC10Api.RecordsetDef AddNew(string source);

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		/// <param name="source">string source</param>
		/// <param name="rowsourceType">optional object rowsourceType</param>
		[CustomMethod]
		[SupportByVersion("OWC10", 1)]
		NetOffice.OWC10Api.RecordsetDef AddNew(string source, object rowsourceType);

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		/// <param name="index">object index</param>
		[SupportByVersion("OWC10", 1)]
		void Delete(object index);

        #endregion

        #region IEnumerable<NetOffice.OWC10Api.RecordsetDef>

        /// <summary>
        /// SupportByVersion OWC10, 1
        /// </summary>
        [SupportByVersion("OWC10", 1)]
        new IEnumerator<NetOffice.OWC10Api.RecordsetDef> GetEnumerator();

        #endregion
    }
}
