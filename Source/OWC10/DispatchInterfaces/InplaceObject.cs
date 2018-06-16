using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.OWC10Api
{
	/// <summary>
	/// DispatchInterface InplaceObject 
	/// SupportByVersion OWC10, 1
	/// </summary>
	[SupportByVersion("OWC10", 1)]
	[EntityType(EntityType.IsDispatchInterface)]
	[TypeId("EDAC92DC-84DB-4CD1-AE98-EB436F4D99EB")]
	public interface InplaceObject : ICOMObject
	{
		#region Methods

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		/// <param name="hwnd">Int32 hwnd</param>
		[SupportByVersion("OWC10", 1)]
		void Activate(Int32 hwnd);

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		void Deactivate();

		#endregion
	}
}
