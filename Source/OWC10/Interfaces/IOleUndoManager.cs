using System.Reflection;
using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.OWC10Api
{
	/// <summary>
	/// Interface IOleUndoManager 
	/// SupportByVersion OWC10, 1
	/// </summary>
	[SupportByVersion("OWC10", 1)]
	[EntityType(EntityType.IsInterface)]
	[TypeId("D001F200-EF97-11CE-9BC9-00AA00608E01")]
	public interface IOleUndoManager : ICOMObject
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
		/// <param name="pdwState">Int32 pdwState</param>
		[SupportByVersion("OWC10", 1)]
		Int32 GetOpenParentState(out Int32 pdwState);

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		/// <param name="pUU">NetOffice.OWC10Api.IOleUndoUnit pUU</param>
		[SupportByVersion("OWC10", 1)]
		Int32 DiscardFrom(NetOffice.OWC10Api.IOleUndoUnit pUU);

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		/// <param name="pUU">NetOffice.OWC10Api.IOleUndoUnit pUU</param>
		[SupportByVersion("OWC10", 1)]
		Int32 UndoTo(NetOffice.OWC10Api.IOleUndoUnit pUU);

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		/// <param name="pUU">NetOffice.OWC10Api.IOleUndoUnit pUU</param>
		[SupportByVersion("OWC10", 1)]
		Int32 RedoTo(NetOffice.OWC10Api.IOleUndoUnit pUU);

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		/// <param name="ppEnum">NetOffice.OWC10Api.IEnumOleUndoUnits ppEnum</param>
		[SupportByVersion("OWC10", 1)]
		Int32 EnumUndoable(out NetOffice.OWC10Api.IEnumOleUndoUnits ppEnum);

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		/// <param name="ppEnum">NetOffice.OWC10Api.IEnumOleUndoUnits ppEnum</param>
		[SupportByVersion("OWC10", 1)]
		Int32 EnumRedoable(out NetOffice.OWC10Api.IEnumOleUndoUnits ppEnum);

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		/// <param name="pbstr">string pbstr</param>
		[SupportByVersion("OWC10", 1)]
		Int32 GetLastUndoDescription(out string pbstr);

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		/// <param name="pbstr">string pbstr</param>
		[SupportByVersion("OWC10", 1)]
		Int32 GetLastRedoDescription(out string pbstr);

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		/// <param name="fEnable">Int32 fEnable</param>
		[SupportByVersion("OWC10", 1)]
		Int32 Enable(Int32 fEnable);

		#endregion
	}
}
