using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.VisioApi
{
	/// <summary>
	/// DispatchInterface IVBUndoManager 
	/// SupportByVersion Visio, 11,12,14,15,16
	/// </summary>
	[SupportByVersion("Visio", 11,12,14,15,16)]
	[EntityType(EntityType.IsDispatchInterface)]
	[TypeId("000D1306-0000-0000-C000-000000000046")]
	public interface IVBUndoManager : ICOMObject
	{
		#region Methods

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="pUnit">NetOffice.VisioApi.IVBUndoUnit pUnit</param>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		void Add(NetOffice.VisioApi.IVBUndoUnit pUnit);

		#endregion
	}
}
