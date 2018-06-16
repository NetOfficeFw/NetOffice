using System.Reflection;
using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.OWC10Api
{
	/// <summary>
	/// DispatchInterface ChErrorBars 
	/// SupportByVersion OWC10, 1
	/// </summary>
	[SupportByVersion("OWC10", 1)]
	[EntityType(EntityType.IsDispatchInterface)]
	[TypeId("F5B39AEB-1480-11D3-8549-00C04FAC67D7")]
	public interface ChErrorBars : ICOMObject
	{
		#region Properties

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get/Set
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		NetOffice.OWC10Api.Enums.ChartEndStyleEnum EndStyle { get; set; }

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get/Set
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		NetOffice.OWC10Api.Enums.ChartErrorBarDirectionEnum Direction { get; set; }

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		Int32 Index { get; }

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		NetOffice.OWC10Api.ChLine Line { get; }

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		NetOffice.OWC10Api.ChSeries Parent { get; }

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get/Set
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		Double Amount { get; set; }

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get/Set
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		NetOffice.OWC10Api.Enums.ChartErrorBarIncludeEnum Include { get; set; }

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get/Set
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		NetOffice.OWC10Api.Enums.ChartErrorBarTypeEnum Type { get; set; }

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		Int32 Top { get; }

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		Int32 Left { get; }

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		Int32 Bottom { get; }

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		Int32 Right { get; }

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		NetOffice.OWC10Api.Enums.ChartSelectionsEnum ObjectType { get; }

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		/// <param name="dimension">NetOffice.OWC10Api.Enums.ChartErrorBarCustomValuesEnum dimension</param>
		/// <param name="dataSourceIndex">Int32 dataSourceIndex</param>
		/// <param name="dataReference">optional object dataReference</param>
		[SupportByVersion("OWC10", 1)]
		void SetData(NetOffice.OWC10Api.Enums.ChartErrorBarCustomValuesEnum dimension, Int32 dataSourceIndex, object dataReference);

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		/// <param name="dimension">NetOffice.OWC10Api.Enums.ChartErrorBarCustomValuesEnum dimension</param>
		/// <param name="dataSourceIndex">Int32 dataSourceIndex</param>
		[CustomMethod]
		[SupportByVersion("OWC10", 1)]
		void SetData(NetOffice.OWC10Api.Enums.ChartErrorBarCustomValuesEnum dimension, Int32 dataSourceIndex);

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		/// <param name="errorBarDirection">NetOffice.OWC10Api.Enums.ChartErrorBarCustomValuesEnum errorBarDirection</param>
		[SupportByVersion("OWC10", 1)]
		string GetDataReference(NetOffice.OWC10Api.Enums.ChartErrorBarCustomValuesEnum errorBarDirection);

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		/// <param name="errorBarDirection">NetOffice.OWC10Api.Enums.ChartErrorBarCustomValuesEnum errorBarDirection</param>
		[SupportByVersion("OWC10", 1)]
		Int32 GetDataSourceIndex(NetOffice.OWC10Api.Enums.ChartErrorBarCustomValuesEnum errorBarDirection);

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		/// <param name="errorBarDirection">NetOffice.OWC10Api.Enums.ChartErrorBarCustomValuesEnum errorBarDirection</param>
		/// <param name="dataSourceIndex">object dataSourceIndex</param>
		/// <param name="dataReference">object dataReference</param>
		[SupportByVersion("OWC10", 1)]
		void GetData(NetOffice.OWC10Api.Enums.ChartErrorBarCustomValuesEnum errorBarDirection, out object dataSourceIndex, out object dataReference);

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		void Select();

		#endregion
	}
}
