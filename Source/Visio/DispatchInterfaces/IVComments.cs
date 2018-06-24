using System.Collections;
using System.Collections.Generic;
using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;
using NetOffice.CollectionsGeneric;

namespace NetOffice.VisioApi
{
	/// <summary>
	/// DispatchInterface IVComments 
	/// SupportByVersion Visio, 15, 16
	/// </summary>
	[SupportByVersion("Visio", 15, 16)]
	[EntityType(EntityType.IsDispatchInterface), BaseType, Enumerator(Enumerator.Reference, EnumeratorInvoke.Property, "Visio", 15, 16), HasIndexProperty(IndexInvoke.Property, "Item")]
	[TypeId("000D0743-0000-0000-C000-000000000046")]
    [CoClassSource(typeof(NetOffice.VisioApi.Comments))]
    public interface IVComments : ICOMObject, IEnumerableProvider<NetOffice.VisioApi.IVComment>
	{
		#region Properties

		/// <summary>
		/// SupportByVersion Visio 15,16
		/// Get
		/// </summary>
		[SupportByVersion("Visio", 15, 16)]
		[BaseResult]
		NetOffice.VisioApi.IVApplication Application { get; }

		/// <summary>
		/// SupportByVersion Visio 15,16
		/// Get
		/// </summary>
		[SupportByVersion("Visio", 15, 16)]
		Int16 Stat { get; }

		/// <summary>
		/// SupportByVersion Visio 15,16
		/// Get
		/// </summary>
		[SupportByVersion("Visio", 15, 16)]
		[BaseResult]
		NetOffice.VisioApi.IVDocument Document { get; }

		/// <summary>
		/// SupportByVersion Visio 15,16
		/// Get
		/// </summary>
		[SupportByVersion("Visio", 15, 16)]
		Int16 ObjectType { get; }

		/// <summary>
		/// SupportByVersion Visio 15,16
		/// Get
		/// </summary>
		[SupportByVersion("Visio", 15, 16)]
		Int32 Count { get; }

		/// <summary>
		/// SupportByVersion Visio 15,16
		/// Get
		/// </summary>
		/// <param name="index">Int32 index</param>
		[SupportByVersion("Visio", 15, 16)]
		[BaseResult]
		[NetRuntimeSystem.Runtime.CompilerServices.IndexerName("Item"), IndexProperty]
		NetOffice.VisioApi.IVComment this[Int32 index] { get; }

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion Visio 15,16
		/// </summary>
		/// <param name="text">string text</param>
		[SupportByVersion("Visio", 15, 16)]
		[BaseResult]
		NetOffice.VisioApi.IVComment Add(string text);

		/// <summary>
		/// SupportByVersion Visio 15,16
		/// </summary>
		[SupportByVersion("Visio", 15, 16)]
		void DeleteAll();

        #endregion

        #region IEnumerable<NetOffice.VisioApi.IVComment>

        /// <summary>
        /// SupportByVersion Visio, 15, 16
        /// </summary>
        [SupportByVersion("Visio", 15, 16)]
        new IEnumerator<NetOffice.VisioApi.IVComment> GetEnumerator();

        #endregion
    }
}
