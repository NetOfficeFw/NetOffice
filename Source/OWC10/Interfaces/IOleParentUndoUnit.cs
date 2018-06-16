using System.Reflection;
using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.OWC10Api
{
	/// <summary>
	/// Interface IOleParentUndoUnit 
	/// SupportByVersion OWC10, 1
	/// </summary>
	[SupportByVersion("OWC10", 1)]
	[EntityType(EntityType.IsInterface)]
	[TypeId("A1FAF330-EF97-11CE-9BC9-00AA00608E01")]
	public interface IOleParentUndoUnit : IOleUndoUnit
	{
		#region Methods

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		/// <param name="pPUU">NetOffice.OWC10Api.IOleParentUndoUnit pPUU</param>
		[SupportByVersion("OWC10", 1)]
		Int32 Open(NetOffice.OWC10Api.IOleParentUndoUnit pPUU);

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		/// <param name="pPUU">NetOffice.OWC10Api.IOleParentUndoUnit pPUU</param>
		/// <param name="fCommit">Int32 fCommit</param>
		[SupportByVersion("OWC10", 1)]
		Int32 Close(NetOffice.OWC10Api.IOleParentUndoUnit pPUU, Int32 fCommit);

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		/// <param name="pUU">NetOffice.OWC10Api.IOleUndoUnit pUU</param>
		[SupportByVersion("OWC10", 1)]
		Int32 Add(NetOffice.OWC10Api.IOleUndoUnit pUU);

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		/// <param name="pUU">NetOffice.OWC10Api.IOleUndoUnit pUU</param>
		[SupportByVersion("OWC10", 1)]
		Int32 FindUnit(NetOffice.OWC10Api.IOleUndoUnit pUU);

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		/// <param name="pdwState">Int32 pdwState</param>
		[SupportByVersion("OWC10", 1)]
		Int32 GetParentState(out Int32 pdwState);

		#endregion
	}
}
