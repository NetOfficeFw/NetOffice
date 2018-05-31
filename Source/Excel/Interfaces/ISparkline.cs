using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.ExcelApi
{
	/// <summary>
	/// Interface ISparkline 
	/// SupportByVersion Excel, 14,15,16
	/// </summary>
	[SupportByVersion("Excel", 14,15,16)]
	[EntityType(EntityType.IsInterface)]
	public interface ISparkline : ICOMObject
	{
		#region Properties

		/// <summary>
		/// SupportByVersion Excel 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Excel", 14,15,16)]
		NetOffice.ExcelApi.Application Application { get; }

		/// <summary>
		/// SupportByVersion Excel 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Excel", 14,15,16)]
		NetOffice.ExcelApi.Enums.XlCreator Creator { get; }

		/// <summary>
		/// SupportByVersion Excel 14, 15, 16
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		[SupportByVersion("Excel", 14,15,16), ProxyResult]
		object Parent { get; }

		/// <summary>
		/// SupportByVersion Excel 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Excel", 14,15,16)]
		NetOffice.ExcelApi.Range Location { get; set; }

		/// <summary>
		/// SupportByVersion Excel 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Excel", 14,15,16)]
		string SourceData { get; set; }

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion Excel 14, 15, 16
		/// </summary>
		/// <param name="range">NetOffice.ExcelApi.Range range</param>
		[SupportByVersion("Excel", 14,15,16)]
		Int32 ModifyLocation(NetOffice.ExcelApi.Range range);

		/// <summary>
		/// SupportByVersion Excel 14, 15, 16
		/// </summary>
		/// <param name="formula">string formula</param>
		[SupportByVersion("Excel", 14,15,16)]
		Int32 ModifySourceData(string formula);

		#endregion
	}
}
