using System.Collections;
using System.Collections.Generic;
using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.OWC10Api
{
	/// <summary>
	/// DispatchInterface SchemaRelationships 
	/// SupportByVersion OWC10, 1
	/// </summary>
	[SupportByVersion("OWC10", 1)]
	[EntityType(EntityType.IsDispatchInterface), HasIndexProperty(IndexInvoke.Property, "Item")]
	[TypeId("F5B39A6C-1480-11D3-8549-00C04FAC67D7")]
	public interface SchemaRelationships : ICOMObject, IEnumerable<NetOffice.OWC10Api.SchemaRelationship>
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
		NetOffice.OWC10Api.SchemaRelationship this[object index] { get; }

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		/// <param name="name">string name</param>
		/// <param name="manySchemaRowsource">string manySchemaRowsource</param>
		/// <param name="oneSchemaRowsource">string oneSchemaRowsource</param>
		/// <param name="manySchemaField">string manySchemaField</param>
		/// <param name="oneSchemaField">string oneSchemaField</param>
		[SupportByVersion("OWC10", 1)]
		NetOffice.OWC10Api.SchemaRelationship Add(string name, string manySchemaRowsource, string oneSchemaRowsource, string manySchemaField, string oneSchemaField);

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		/// <param name="name">string name</param>
		/// <param name="manySchemaRowsource">string manySchemaRowsource</param>
		/// <param name="oneSchemaRowsource">string oneSchemaRowsource</param>
		/// <param name="manySchemaField">string manySchemaField</param>
		/// <param name="oneSchemaField">string oneSchemaField</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("OWC10", 1)]
		NetOffice.OWC10Api.SchemaRelationship AddNew(string name, string manySchemaRowsource, string oneSchemaRowsource, string manySchemaField, string oneSchemaField);

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		/// <param name="index">object index</param>
		[SupportByVersion("OWC10", 1)]
		void Delete(object index);

		#endregion
	}
}
