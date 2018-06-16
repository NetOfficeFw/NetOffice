using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.OWC10Api
{
	/// <summary>
	/// DispatchInterface SelectionHighlight 
	/// SupportByVersion OWC10, 1
	/// </summary>
	[SupportByVersion("OWC10", 1)]
	[EntityType(EntityType.IsDispatchInterface)]
	[TypeId("58573A80-5025-11D3-BE84-0050041DB15A")]
	public interface SelectionHighlight : ICOMObject
	{
		#region Methods

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		/// <param name="viewSurface">NetOffice.OWC10Api.ViewSurface viewSurface</param>
		[SupportByVersion("OWC10", 1)]
		void Highlight(NetOffice.OWC10Api.ViewSurface viewSurface);

		#endregion
	}
}
