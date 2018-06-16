using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.OWC10Api
{
	/// <summary>
	/// DispatchInterface PropertyNotify 
	/// SupportByVersion OWC10, 1
	/// </summary>
	[SupportByVersion("OWC10", 1)]
	[EntityType(EntityType.IsDispatchInterface)]
	[TypeId("7BB4EDA1-862A-4AB2-92F2-557E1BAB3408")]
	public interface PropertyNotify : ICOMObject
	{
		#region Methods

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		/// <param name="_object">object object</param>
		/// <param name="dispid">Int32 dispid</param>
		[SupportByVersion("OWC10", 1)]
		void OnPropertyChange(object _object, Int32 dispid);

		#endregion
	}
}
