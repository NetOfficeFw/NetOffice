using System;
using NetOffice;
using NetOffice.Attributes;
namespace NetOffice.ExcelApi.Enums
{
	 /// <summary>
	 /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
	 /// </summary>
	 ///<remarks> MSDN Online Documentation: <see href="https://docs.microsoft.com/en-us/office/vba/api/Excel.XlSummaryReportType"/> </remarks>
	[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
	[EntityType(EntityType.IsEnum)]
	public enum XlSummaryReportType
	{
		 /// <summary>
		 /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>-4148</remarks>
		 [SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		 xlSummaryPivotTable = -4148,

		 /// <summary>
		 /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>1</remarks>
		 [SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		 xlStandardSummary = 1
	}
}