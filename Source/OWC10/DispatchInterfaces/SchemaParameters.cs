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
	/// DispatchInterface SchemaParameters 
	/// SupportByVersion OWC10, 1
	/// </summary>
	[SupportByVersion("OWC10", 1)]
	[EntityType(EntityType.IsDispatchInterface), Enumerator(Enumerator.Reference, EnumeratorInvoke.Property, "OWC10", 1), HasIndexProperty(IndexInvoke.Property, "Item")]
	[TypeId("F5B39AED-1480-11D3-8549-00C04FAC67D7")]
	public interface SchemaParameters : ICOMObject, IEnumerableProvider<NetOffice.OWC10Api.SchemaParameter>
	{
		#region Properties

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get
		/// </summary>
		/// <param name="index">object index</param>
		[SupportByVersion("OWC10", 1)]
		[NetRuntimeSystem.Runtime.CompilerServices.IndexerName("Item"), IndexProperty]
		NetOffice.OWC10Api.SchemaParameter this[object index] { get; }

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
		/// <param name="name">string name</param>
		/// <param name="dataType">optional object dataType</param>
		/// <param name="size">optional object size</param>
		/// <param name="scale">optional object scale</param>
		/// <param name="precision">optional object precision</param>
		/// <param name="direction">optional object direction</param>
		[SupportByVersion("OWC10", 1)]
		NetOffice.OWC10Api.SchemaParameter Add(string name, object dataType, object size, object scale, object precision, object direction);

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		/// <param name="name">string name</param>
		[CustomMethod]
		[SupportByVersion("OWC10", 1)]
		NetOffice.OWC10Api.SchemaParameter Add(string name);

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		/// <param name="name">string name</param>
		/// <param name="dataType">optional object dataType</param>
		[CustomMethod]
		[SupportByVersion("OWC10", 1)]
		NetOffice.OWC10Api.SchemaParameter Add(string name, object dataType);

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		/// <param name="name">string name</param>
		/// <param name="dataType">optional object dataType</param>
		/// <param name="size">optional object size</param>
		[CustomMethod]
		[SupportByVersion("OWC10", 1)]
		NetOffice.OWC10Api.SchemaParameter Add(string name, object dataType, object size);

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		/// <param name="name">string name</param>
		/// <param name="dataType">optional object dataType</param>
		/// <param name="size">optional object size</param>
		/// <param name="scale">optional object scale</param>
		[CustomMethod]
		[SupportByVersion("OWC10", 1)]
		NetOffice.OWC10Api.SchemaParameter Add(string name, object dataType, object size, object scale);

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		/// <param name="name">string name</param>
		/// <param name="dataType">optional object dataType</param>
		/// <param name="size">optional object size</param>
		/// <param name="scale">optional object scale</param>
		/// <param name="precision">optional object precision</param>
		[CustomMethod]
		[SupportByVersion("OWC10", 1)]
		NetOffice.OWC10Api.SchemaParameter Add(string name, object dataType, object size, object scale, object precision);

        #endregion

        #region IEnumerable<NetOffice.OWC10Api.SchemaParameter>

        /// <summary>
        /// SupportByVersion OWC10, 1
        /// </summary>
        [SupportByVersion("OWC10", 1)]
        new IEnumerator<NetOffice.OWC10Api.SchemaParameter> GetEnumerator();

        #endregion
    }
}
