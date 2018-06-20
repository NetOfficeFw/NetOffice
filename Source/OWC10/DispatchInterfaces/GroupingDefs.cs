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
	/// DispatchInterface GroupingDefs 
	/// SupportByVersion OWC10, 1
	/// </summary>
	[SupportByVersion("OWC10", 1)]
	[EntityType(EntityType.IsDispatchInterface), Enumerator(Enumerator.Reference, EnumeratorInvoke.Property, "OWC10", 1), HasIndexProperty(IndexInvoke.Property, "Item")]
	[TypeId("F5B39A7C-1480-11D3-8549-00C04FAC67D7")]
	public interface GroupingDefs : ICOMObject, IEnumerableProvider<NetOffice.OWC10Api.GroupingDef>
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
		NetOffice.OWC10Api.GroupingDef this[object index] { get; }

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		/// <param name="groupingDefName">string groupingDefName</param>
		/// <param name="groupingFieldName">string groupingFieldName</param>
		/// <param name="pageFieldName">string pageFieldName</param>
		/// <param name="index">optional object index</param>
		[SupportByVersion("OWC10", 1)]
		NetOffice.OWC10Api.GroupingDef Add(string groupingDefName, string groupingFieldName, string pageFieldName, object index);

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		/// <param name="groupingDefName">string groupingDefName</param>
		/// <param name="groupingFieldName">string groupingFieldName</param>
		/// <param name="pageFieldName">string pageFieldName</param>
		[CustomMethod]
		[SupportByVersion("OWC10", 1)]
		NetOffice.OWC10Api.GroupingDef Add(string groupingDefName, string groupingFieldName, string pageFieldName);

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		/// <param name="groupingDefName">string groupingDefName</param>
		/// <param name="groupingFieldName">string groupingFieldName</param>
		/// <param name="pageFieldName">string pageFieldName</param>
		/// <param name="totalType">NetOffice.OWC10Api.Enums.DscTotalTypeEnum totalType</param>
		/// <param name="index">optional object index</param>
		[SupportByVersion("OWC10", 1)]
		NetOffice.OWC10Api.GroupingDef AddTotal(string groupingDefName, string groupingFieldName, string pageFieldName, NetOffice.OWC10Api.Enums.DscTotalTypeEnum totalType, object index);

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		/// <param name="groupingDefName">string groupingDefName</param>
		/// <param name="groupingFieldName">string groupingFieldName</param>
		/// <param name="pageFieldName">string pageFieldName</param>
		/// <param name="totalType">NetOffice.OWC10Api.Enums.DscTotalTypeEnum totalType</param>
		[CustomMethod]
		[SupportByVersion("OWC10", 1)]
		NetOffice.OWC10Api.GroupingDef AddTotal(string groupingDefName, string groupingFieldName, string pageFieldName, NetOffice.OWC10Api.Enums.DscTotalTypeEnum totalType);

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		/// <param name="index">object index</param>
		[SupportByVersion("OWC10", 1)]
		void Delete(object index);

        #endregion

        #region IEnumerable<NetOffice.OWC10Api.GroupingDef>

        /// <summary>
        /// SupportByVersion OWC10, 1
        /// </summary>
        [SupportByVersion("OWC10", 1)]
        new IEnumerator<NetOffice.OWC10Api.GroupingDef> GetEnumerator();

        #endregion
    }
}
