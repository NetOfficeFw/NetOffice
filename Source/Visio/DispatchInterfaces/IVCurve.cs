using System.Reflection;
using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.VisioApi
{
	/// <summary>
	/// DispatchInterface IVCurve 
	/// SupportByVersion Visio, 11,12,14,15,16
	/// </summary>
	[SupportByVersion("Visio", 11,12,14,15,16)]
	[EntityType(EntityType.IsDispatchInterface), BaseType]
	[TypeId("000D0722-0000-0000-C000-000000000046")]
    [CoClassSource(typeof(NetOffice.VisioApi.Curve))]
    public interface IVCurve : ICOMObject
	{
		#region Properties

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		[BaseResult]
		NetOffice.VisioApi.IVApplication Application { get; }

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		Int16 ObjectType { get; }

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		Int16 Closed { get; }

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		Double Start { get; }

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		Double End { get; }

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="tolerance">Double tolerance</param>
		/// <param name="xyArray">Double[] xyArray</param>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		void Points(Double tolerance, out Double[] xyArray);

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="t">Double t</param>
		/// <param name="x">Double x</param>
		/// <param name="y">Double y</param>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		void Point(Double t, out Double x, out Double y);

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="t">Double t</param>
		/// <param name="n">Int16 n</param>
		/// <param name="x">Double x</param>
		/// <param name="y">Double y</param>
		/// <param name="dxdt">Double dxdt</param>
		/// <param name="dydt">Double dydt</param>
		/// <param name="ddxdt">Double ddxdt</param>
		/// <param name="ddydt">Double ddydt</param>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		void PointAndDerivatives(Double t, Int16 n, out Double x, out Double y, out Double dxdt, out Double dydt, out Double ddxdt, out Double ddydt);

		#endregion
	}
}
