using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.ExcelApi
{
	/// <summary>
	/// Interface IQuickAnalysis 
	/// SupportByVersion Excel, 15, 16
	/// </summary>
	[SupportByVersion("Excel", 15, 16)]
	[EntityType(EntityType.IsInterface)]
	[TypeId("000244D0-0001-0000-C000-000000000046")]
	public interface IQuickAnalysis : ICOMObject
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

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion Excel 15,16
		/// </summary>
		/// <param name="xlQuickAnalysisMode">optional NetOffice.ExcelApi.Enums.XlQuickAnalysisMode XlQuickAnalysisMode = 0</param>
		[SupportByVersion("Excel", 15, 16)]
		Int32 Show(object xlQuickAnalysisMode);

		/// <summary>
		/// SupportByVersion Excel 15,16
		/// </summary>
		[CustomMethod]
		[SupportByVersion("Excel", 15, 16)]
		Int32 Show();

		/// <summary>
		/// SupportByVersion Excel 15,16
		/// </summary>
		/// <param name="xlQuickAnalysisMode">optional NetOffice.ExcelApi.Enums.XlQuickAnalysisMode XlQuickAnalysisMode = 0</param>
		[SupportByVersion("Excel", 15, 16)]
		Int32 Hide(object xlQuickAnalysisMode);

		/// <summary>
		/// SupportByVersion Excel 15,16
		/// </summary>
		[CustomMethod]
		[SupportByVersion("Excel", 15, 16)]
		Int32 Hide();

		#endregion
	}
}
