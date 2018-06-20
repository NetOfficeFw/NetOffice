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
	/// DispatchInterface SchemaFields 
	/// SupportByVersion OWC10, 1
	/// </summary>
	[SupportByVersion("OWC10", 1)]
	[EntityType(EntityType.IsDispatchInterface), Enumerator(Enumerator.Reference, EnumeratorInvoke.Property, "OWC10", 1), HasIndexProperty(IndexInvoke.Property, "Item")]
	[TypeId("F5B39AA3-1480-11D3-8549-00C04FAC67D7")]
	public interface SchemaFields : ICOMObject, IEnumerableProvider<NetOffice.OWC10Api.SchemaField>
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
		[BaseResult]
		[NetRuntimeSystem.Runtime.CompilerServices.IndexerName("Item"), IndexProperty]
		NetOffice.OWC10Api.SchemaField this[object index] { get; }

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		/// <param name="name">string name</param>
		/// <param name="dataType">NetOffice.ADODBApi.Enums.DataTypeEnum dataType</param>
		/// <param name="length">optional object length</param>
		[SupportByVersion("OWC10", 1)]
		[BaseResult]
		NetOffice.OWC10Api.SchemaField Add(string name, NetOffice.ADODBApi.Enums.DataTypeEnum dataType, object length);

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		/// <param name="name">string name</param>
		/// <param name="dataType">NetOffice.ADODBApi.Enums.DataTypeEnum dataType</param>
		[CustomMethod]
		[BaseResult]
		[SupportByVersion("OWC10", 1)]
		NetOffice.OWC10Api.SchemaField Add(string name, NetOffice.ADODBApi.Enums.DataTypeEnum dataType);

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		/// <param name="index">object index</param>
		[SupportByVersion("OWC10", 1)]
		void Delete(object index);

        #endregion

        #region IEnumerable<NetOffice.OWC10Api.SchemaField>

        /// <summary>
        /// SupportByVersion OWC10, 1
        /// </summary>
        [SupportByVersion("OWC10", 1)]
        new IEnumerator<NetOffice.OWC10Api.SchemaField> GetEnumerator();

        #endregion
    }
}
