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
	/// DispatchInterface IVValidationRuleSets 
	/// SupportByVersion Visio, 14,15,16
	/// </summary>
	[SupportByVersion("Visio", 14,15,16)]
	[EntityType(EntityType.IsDispatchInterface), BaseType, Enumerator(Enumerator.Reference, EnumeratorInvoke.Property, "Visio", 14, 15, 16), HasIndexProperty(IndexInvoke.Property, "Item")]
	[TypeId("000D073B-0000-0000-C000-000000000046")]
    [CoClassSource(typeof(NetOffice.VisioApi.ValidationRuleSets))]
    public interface IVValidationRuleSets : ICOMObject, IEnumerableProvider<NetOffice.VisioApi.IVValidationRuleSet>
	{
		#region Properties

		/// <summary>
		/// SupportByVersion Visio 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Visio", 14,15,16)]
		[BaseResult]
		NetOffice.VisioApi.IVApplication Application { get; }

		/// <summary>
		/// SupportByVersion Visio 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Visio", 14,15,16)]
		Int16 Stat { get; }

		/// <summary>
		/// SupportByVersion Visio 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Visio", 14,15,16)]
		[BaseResult]
		NetOffice.VisioApi.IVDocument Document { get; }

		/// <summary>
		/// SupportByVersion Visio 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Visio", 14,15,16)]
		Int16 ObjectType { get; }

		/// <summary>
		/// SupportByVersion Visio 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Visio", 14,15,16)]
		Int32 Count { get; }

		/// <summary>
		/// SupportByVersion Visio 14, 15, 16
		/// Get
		/// </summary>
		/// <param name="nameUOrIndex">object nameUOrIndex</param>
		[SupportByVersion("Visio", 14,15,16)]
		[BaseResult]
		[NetRuntimeSystem.Runtime.CompilerServices.IndexerName("Item"), IndexProperty]
		NetOffice.VisioApi.IVValidationRuleSet this[object nameUOrIndex] { get; }

		/// <summary>
		/// SupportByVersion Visio 14, 15, 16
		/// Get
		/// </summary>
		/// <param name="ruleID">Int32 ruleID</param>
		[SupportByVersion("Visio", 14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		NetOffice.VisioApi.IVValidationRuleSet get_ItemFromID(Int32 ruleID);

		/// <summary>
		/// SupportByVersion Visio 14, 15, 16
		/// Alias for get_ItemFromID
		/// </summary>
		/// <param name="ruleID">Int32 ruleID</param>
		[SupportByVersion("Visio", 14,15,16), Redirect("get_ItemFromID")]
		NetOffice.VisioApi.IVValidationRuleSet ItemFromID(Int32 ruleID);

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion Visio 14, 15, 16
		/// </summary>
		/// <param name="nameU">string nameU</param>
		[SupportByVersion("Visio", 14,15,16)]
		[BaseResult]
		NetOffice.VisioApi.IVValidationRuleSet Add(string nameU);

		/// <summary>
		/// SupportByVersion Visio 14, 15, 16
		/// </summary>
		/// <param name="ruleSet">NetOffice.VisioApi.IVValidationRuleSet ruleSet</param>
		/// <param name="nameU">optional string NameU = </param>
		[SupportByVersion("Visio", 14,15,16)]
		[BaseResult]
		NetOffice.VisioApi.IVValidationRuleSet AddCopy(NetOffice.VisioApi.IVValidationRuleSet ruleSet, object nameU);

		/// <summary>
		/// SupportByVersion Visio 14, 15, 16
		/// </summary>
		/// <param name="ruleSet">NetOffice.VisioApi.IVValidationRuleSet ruleSet</param>
		[CustomMethod]
		[BaseResult]
		[SupportByVersion("Visio", 14,15,16)]
		NetOffice.VisioApi.IVValidationRuleSet AddCopy(NetOffice.VisioApi.IVValidationRuleSet ruleSet);

        #endregion

        #region IEnumerable<NetOffice.VisioApi.IVValidationRuleSet>

        /// <summary>
        /// SupportByVersion Visio, 14,15,16
        /// </summary>
        [SupportByVersion("Visio", 14, 15, 16)]
        new IEnumerator<NetOffice.VisioApi.IVValidationRuleSet> GetEnumerator();

        #endregion
    }
}
