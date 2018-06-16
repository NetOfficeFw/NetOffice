using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.OWC10Api
{
	/// <summary>
	/// Interface ISimpleDataConverter 
	/// SupportByVersion OWC10, 1
	/// </summary>
	[SupportByVersion("OWC10", 1)]
	[EntityType(EntityType.IsInterface)]
	[TypeId("78667670-3C3D-11D2-91F9-006097C97F9B")]
	public interface ISimpleDataConverter : ICOMObject
	{
		#region Methods

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		/// <param name="varSrc">object varSrc</param>
		/// <param name="vtDest">Int32 vtDest</param>
		/// <param name="pUnknownElement">object pUnknownElement</param>
		/// <param name="pvarDest">object pvarDest</param>
		[SupportByVersion("OWC10", 1)]
		Int32 ConvertData(object varSrc, Int32 vtDest, object pUnknownElement, object pvarDest);

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		/// <param name="vt1">Int32 vt1</param>
		/// <param name="vt2">Int32 vt2</param>
		[SupportByVersion("OWC10", 1)]
		Int32 CanConvertData(Int32 vt1, Int32 vt2);

		#endregion
	}
}
