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
	/// DispatchInterface IVValidationIssues 
	/// SupportByVersion Visio, 14,15,16
	/// </summary>
	[SupportByVersion("Visio", 14,15,16)]
	[EntityType(EntityType.IsDispatchInterface), BaseType, Enumerator(Enumerator.Reference, EnumeratorInvoke.Property, "Visio", 14, 15, 16), HasIndexProperty(IndexInvoke.Property, "Item")]
	[TypeId("000D073F-0000-0000-C000-000000000046")]
    [CoClassSource(typeof(NetOffice.VisioApi.ValidationIssues))]
    public interface IVValidationIssues : ICOMObject, IEnumerableProvider<NetOffice.VisioApi.IVValidationIssue>
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
		/// <param name="index">Int32 index</param>
		[SupportByVersion("Visio", 14,15,16)]
		[BaseResult]
		[NetRuntimeSystem.Runtime.CompilerServices.IndexerName("Item"), IndexProperty]
		NetOffice.VisioApi.IVValidationIssue this[Int32 index] { get; }

		/// <summary>
		/// SupportByVersion Visio 14, 15, 16
		/// Get
		/// </summary>
		/// <param name="issueID">Int32 issueID</param>
		[SupportByVersion("Visio", 14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		NetOffice.VisioApi.IVValidationIssue get_ItemFromID(Int32 issueID);

		/// <summary>
		/// SupportByVersion Visio 14, 15, 16
		/// Alias for get_ItemFromID
		/// </summary>
		/// <param name="issueID">Int32 issueID</param>
		[SupportByVersion("Visio", 14,15,16), Redirect("get_ItemFromID")]
		NetOffice.VisioApi.IVValidationIssue ItemFromID(Int32 issueID);

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion Visio 14, 15, 16
		/// </summary>
		[SupportByVersion("Visio", 14,15,16)]
		void Clear();

        #endregion

        #region IEnumerable<NetOffice.VisioApi.IVValidationIssue>

        /// <summary>
        /// SupportByVersion Visio, 14,15,16
        /// </summary>
        [SupportByVersion("Visio", 14, 15, 16)]
        new IEnumerator<NetOffice.VisioApi.IVValidationIssue> GetEnumerator();

        #endregion
    }
}
