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
	/// DispatchInterface PageFields 
	/// SupportByVersion OWC10, 1
	/// </summary>
	[SupportByVersion("OWC10", 1)]
	[EntityType(EntityType.IsDispatchInterface), Enumerator(Enumerator.Reference, EnumeratorInvoke.Property, "OWC10", 1), HasIndexProperty(IndexInvoke.Property, "Item")]
	[TypeId("F5B39AB1-1480-11D3-8549-00C04FAC67D7")]
	public interface PageFields : ICOMObject, IEnumerableProvider<NetOffice.OWC10Api.PageField>
	{
		#region Properties

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get
		/// </summary>
		/// <param name="index">object index</param>
		[SupportByVersion("OWC10", 1)]
		[NetRuntimeSystem.Runtime.CompilerServices.IndexerName("Item"), IndexProperty]
		NetOffice.OWC10Api.PageField this[object index] { get; }

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
		/// <param name="index">object index</param>
		[SupportByVersion("OWC10", 1)]
		void Delete(object index);

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		/// <param name="source">object source</param>
		/// <param name="fieldType">optional object fieldType</param>
		/// <param name="name">optional object name</param>
		/// <param name="totalType">optional NetOffice.OWC10Api.Enums.DscTotalTypeEnum TotalType = 0</param>
		/// <param name="index">optional object index</param>
		[SupportByVersion("OWC10", 1)]
		NetOffice.OWC10Api.PageField Add(object source, object fieldType, object name, object totalType, object index);

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		/// <param name="source">object source</param>
		[CustomMethod]
		[SupportByVersion("OWC10", 1)]
		NetOffice.OWC10Api.PageField Add(object source);

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		/// <param name="source">object source</param>
		/// <param name="fieldType">optional object fieldType</param>
		[CustomMethod]
		[SupportByVersion("OWC10", 1)]
		NetOffice.OWC10Api.PageField Add(object source, object fieldType);

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		/// <param name="source">object source</param>
		/// <param name="fieldType">optional object fieldType</param>
		/// <param name="name">optional object name</param>
		[CustomMethod]
		[SupportByVersion("OWC10", 1)]
		NetOffice.OWC10Api.PageField Add(object source, object fieldType, object name);

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		/// <param name="source">object source</param>
		/// <param name="fieldType">optional object fieldType</param>
		/// <param name="name">optional object name</param>
		/// <param name="totalType">optional NetOffice.OWC10Api.Enums.DscTotalTypeEnum TotalType = 0</param>
		[CustomMethod]
		[SupportByVersion("OWC10", 1)]
		NetOffice.OWC10Api.PageField Add(object source, object fieldType, object name, object totalType);

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		/// <param name="source">object source</param>
		/// <param name="fieldType">optional object fieldType</param>
		/// <param name="name">optional object name</param>
		/// <param name="totalType">optional NetOffice.OWC10Api.Enums.DscTotalTypeEnum TotalType = 0</param>
		/// <param name="index">optional object index</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("OWC10", 1)]
		NetOffice.OWC10Api.PageField AddBroken(object source, object fieldType, object name, object totalType, object index);

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		/// <param name="source">object source</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("OWC10", 1)]
		NetOffice.OWC10Api.PageField AddBroken(object source);

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		/// <param name="source">object source</param>
		/// <param name="fieldType">optional object fieldType</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("OWC10", 1)]
		NetOffice.OWC10Api.PageField AddBroken(object source, object fieldType);

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		/// <param name="source">object source</param>
		/// <param name="fieldType">optional object fieldType</param>
		/// <param name="name">optional object name</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("OWC10", 1)]
		NetOffice.OWC10Api.PageField AddBroken(object source, object fieldType, object name);

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		/// <param name="source">object source</param>
		/// <param name="fieldType">optional object fieldType</param>
		/// <param name="name">optional object name</param>
		/// <param name="totalType">optional NetOffice.OWC10Api.Enums.DscTotalTypeEnum TotalType = 0</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("OWC10", 1)]
		NetOffice.OWC10Api.PageField AddBroken(object source, object fieldType, object name, object totalType);

        #endregion

        #region IEnumerable<NetOffice.OWC10Api.PageField>

        /// <summary>
        /// SupportByVersion OWC10, 1
        /// </summary>
        [SupportByVersion("OWC10", 1)]
        new IEnumerator<NetOffice.OWC10Api.PageField> GetEnumerator();

        #endregion
    }
}
