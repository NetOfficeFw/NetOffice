using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.OWC10Api
{
	/// <summary>
	/// DispatchInterface IPivotCopy 
	/// SupportByVersion OWC10, 1
	/// </summary>
	[SupportByVersion("OWC10", 1)]
	[EntityType(EntityType.IsDispatchInterface)]
	[TypeId("CD44E547-FEC9-4ADC-AB6A-3129B44801BA")]
	public interface IPivotCopy : ICOMObject
	{
		#region Methods

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		/// <param name="format">Int32 format</param>
		/// <param name="output">optional string Output = 0</param>
		[SupportByVersion("OWC10", 1)]
		void Render(Int32 format, object output);

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		/// <param name="format">Int32 format</param>
		[CustomMethod]
		[SupportByVersion("OWC10", 1)]
		void Render(Int32 format);

		#endregion
	}
}
