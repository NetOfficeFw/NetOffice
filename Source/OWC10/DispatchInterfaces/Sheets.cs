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
	/// DispatchInterface Sheets 
	/// SupportByVersion OWC10, 1
	/// </summary>
	[SupportByVersion("OWC10", 1)]
	[EntityType(EntityType.IsDispatchInterface), Enumerator(Enumerator.Reference, EnumeratorInvoke.Property, "OWC10", 1), HasIndexProperty(IndexInvoke.Property, "Item")]
	[TypeId("F5B39BA7-1480-11D3-8549-00C04FAC67D7")]
	public interface Sheets : ICOMObject, IEnumerableProvider<object>
	{
		#region Properties

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		[BaseResult]
		NetOffice.OWC10Api.ISpreadsheet Application { get; }

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		Int32 Count { get; }

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		/// <param name="index">object index</param>
		[SupportByVersion("OWC10", 1), ProxyResult]
		[NetRuntimeSystem.Runtime.CompilerServices.IndexerName("Item"), IndexProperty]
		object this[object index] { get; }

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		NetOffice.OWC10Api.Workbook Parent { get; }

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get/Set
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		object Visible { get; set; }

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		/// <param name="before">optional object before</param>
		/// <param name="after">optional object after</param>
		/// <param name="count">optional object count</param>
		/// <param name="type">optional object type</param>
		[SupportByVersion("OWC10", 1)]
		object Add(object before, object after, object count, object type);

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		[CustomMethod]
		[SupportByVersion("OWC10", 1)]
		object Add();

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		/// <param name="before">optional object before</param>
		[CustomMethod]
		[SupportByVersion("OWC10", 1)]
		object Add(object before);

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		/// <param name="before">optional object before</param>
		/// <param name="after">optional object after</param>
		[CustomMethod]
		[SupportByVersion("OWC10", 1)]
		object Add(object before, object after);

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		/// <param name="before">optional object before</param>
		/// <param name="after">optional object after</param>
		/// <param name="count">optional object count</param>
		[CustomMethod]
		[SupportByVersion("OWC10", 1)]
		object Add(object before, object after, object count);

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		/// <param name="before">optional object before</param>
		/// <param name="after">optional object after</param>
		[SupportByVersion("OWC10", 1)]
		void Copy(object before, object after);

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		[CustomMethod]
		[SupportByVersion("OWC10", 1)]
		void Copy();

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		/// <param name="before">optional object before</param>
		[CustomMethod]
		[SupportByVersion("OWC10", 1)]
		void Copy(object before);

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		void Delete();

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		/// <param name="before">optional object before</param>
		/// <param name="after">optional object after</param>
		[SupportByVersion("OWC10", 1)]
		void Move(object before, object after);

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		[CustomMethod]
		[SupportByVersion("OWC10", 1)]
		void Move();

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		/// <param name="before">optional object before</param>
		[CustomMethod]
		[SupportByVersion("OWC10", 1)]
		void Move(object before);

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		/// <param name="replace">optional object replace</param>
		[SupportByVersion("OWC10", 1)]
		void Select(object replace);

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		[CustomMethod]
		[SupportByVersion("OWC10", 1)]
		void Select();

        #endregion

        #region IEnumerable<object>

        /// <summary>
        /// SupportByVersion OWC10, 1
        /// </summary>
        [SupportByVersion("OWC10", 1)]
        new IEnumerator<object> GetEnumerator();

        #endregion
    }
}
