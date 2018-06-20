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
	/// DispatchInterface SchemaRelatedFields 
	/// SupportByVersion OWC10, 1
	/// </summary>
	[SupportByVersion("OWC10", 1)]
	[EntityType(EntityType.IsDispatchInterface), Enumerator(Enumerator.Reference, EnumeratorInvoke.Property, "OWC10", 1), HasIndexProperty(IndexInvoke.Property, "Item")]
	[TypeId("F5B39A6D-1480-11D3-8549-00C04FAC67D7")]
	public interface SchemaRelatedFields : ICOMObject, IEnumerableProvider<NetOffice.OWC10Api.SchemaRelatedField>
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
		NetOffice.OWC10Api.SchemaRelatedField this[object index] { get; }

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		/// <param name="manySchemaField">string manySchemaField</param>
		/// <param name="oneSchemaField">string oneSchemaField</param>
		[SupportByVersion("OWC10", 1)]
		NetOffice.OWC10Api.SchemaRelatedField Add(string manySchemaField, string oneSchemaField);

        #endregion

        #region IEnumerable<NetOffice.OWC10Api.SchemaRelatedField>

        /// <summary>
        /// SupportByVersion OWC10, 1
        /// </summary>
        [SupportByVersion("OWC10", 1)]
        new IEnumerator<NetOffice.OWC10Api.SchemaRelatedField> GetEnumerator();

        #endregion
    }
}
