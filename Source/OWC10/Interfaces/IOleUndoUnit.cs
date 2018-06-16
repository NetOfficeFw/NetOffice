using System.Reflection;
using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.OWC10Api
{
	/// <summary>
	/// Interface IOleUndoUnit 
	/// SupportByVersion OWC10, 1
	/// </summary>
	[SupportByVersion("OWC10", 1)]
	[EntityType(EntityType.IsInterface), BaseType]
	[TypeId("894AD3B0-EF97-11CE-9BC9-00AA00608E01")]
	public interface IOleUndoUnit : ICOMObject
	{
		#region Methods

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		/// <param name="pUndoManager">NetOffice.OWC10Api.IOleUndoManager pUndoManager</param>
		[SupportByVersion("OWC10", 1)]
		Int32 Do(NetOffice.OWC10Api.IOleUndoManager pUndoManager);

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		/// <param name="pbstr">string pbstr</param>
		[SupportByVersion("OWC10", 1)]
		Int32 GetDescription(out string pbstr);

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		/// <param name="pClsid">Guid pClsid</param>
		/// <param name="plID">Int32 plID</param>
		[SupportByVersion("OWC10", 1)]
		Int32 GetUnitType(out Guid pClsid, out Int32 plID);

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		Int32 OnNextAdd();

		#endregion
	}
}
