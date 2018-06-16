using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.OWC10Api
{
	/// <summary>
	/// DispatchInterface DropSource 
	/// SupportByVersion OWC10, 1
	/// </summary>
	[SupportByVersion("OWC10", 1)]
	[EntityType(EntityType.IsDispatchInterface)]
	[TypeId("D6CE4620-E224-11D2-8F35-00600893B533")]
	public interface DropSource : ICOMObject
	{
		#region Methods

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		/// <param name="dwEffect">Int32 dwEffect</param>
		[SupportByVersion("OWC10", 1)]
		void GiveFeedback(Int32 dwEffect);

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		/// <param name="escapePressed">bool escapePressed</param>
		/// <param name="keyState">Int32 keyState</param>
		[SupportByVersion("OWC10", 1)]
		void QueryContinueDrag(bool escapePressed, Int32 keyState);

		#endregion
	}
}
