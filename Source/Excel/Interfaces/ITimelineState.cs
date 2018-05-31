using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.ExcelApi
{
	/// <summary>
	/// Interface ITimelineState 
	/// SupportByVersion Excel, 15, 16
	/// </summary>
	[SupportByVersion("Excel", 15, 16)]
	[EntityType(EntityType.IsInterface)]
	public interface ITimelineState : ICOMObject
	{
		#region Properties

		/// <summary>
		/// SupportByVersion Excel 15,16
		/// Get
		/// </summary>
		[SupportByVersion("Excel", 15, 16)]
		NetOffice.ExcelApi.Application Application { get; }

		/// <summary>
		/// SupportByVersion Excel 15,16
		/// Get
		/// </summary>
		[SupportByVersion("Excel", 15, 16)]
		NetOffice.ExcelApi.Enums.XlCreator Creator { get; }

		/// <summary>
		/// SupportByVersion Excel 15,16
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		[SupportByVersion("Excel", 15, 16), ProxyResult]
		object Parent { get; }

		/// <summary>
		/// SupportByVersion Excel 15,16
		/// Get
		/// </summary>
		[SupportByVersion("Excel", 15, 16)]
		object StartDate { get; }

		/// <summary>
		/// SupportByVersion Excel 15,16
		/// Get
		/// </summary>
		[SupportByVersion("Excel", 15, 16)]
		object EndDate { get; }

		/// <summary>
		/// SupportByVersion Excel 15,16
		/// Get
		/// </summary>
		[SupportByVersion("Excel", 15, 16)]
		NetOffice.ExcelApi.Enums.XlPivotFilterType FilterType { get; }

		/// <summary>
		/// SupportByVersion Excel 15,16
		/// Get
		/// </summary>
		[SupportByVersion("Excel", 15, 16)]
		object FilterValue1 { get; }

		/// <summary>
		/// SupportByVersion Excel 15,16
		/// Get
		/// </summary>
		[SupportByVersion("Excel", 15, 16)]
		object FilterValue2 { get; }

		/// <summary>
		/// SupportByVersion Excel 15,16
		/// Get
		/// </summary>
		[SupportByVersion("Excel", 15, 16)]
		bool SingleRangeFilterState { get; }

		/// <summary>
		/// SupportByVersion Excel 15,16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Excel", 15, 16)]
		NetOffice.ExcelApi.Enums.XlTimeMoving MovingPeriod { get; set; }

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion Excel 15,16
		/// </summary>
		/// <param name="startDate">object startDate</param>
		/// <param name="endDate">object endDate</param>
		[SupportByVersion("Excel", 15, 16)]
		NetOffice.ExcelApi.Enums.XlFilterStatus SetFilterDateRange(object startDate, object endDate);

		#endregion
	}
}
