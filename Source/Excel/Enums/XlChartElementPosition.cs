using System;
using NetOffice;
using NetOffice.Attributes;
namespace NetOffice.ExcelApi.Enums
{
	 /// <summary>
	 /// SupportByVersion Excel 12, 14, 15, 16
	 /// </summary>
	 ///<remarks> MSDN Online Documentation: <see href="https://docs.microsoft.com/en-us/office/vba/api/Excel.XlChartElementPosition"/> </remarks>
	[SupportByVersion("Excel", 12,14,15,16)]
	[EntityType(EntityType.IsEnum)]
	public enum XlChartElementPosition
	{
		 /// <summary>
		 /// SupportByVersion Excel 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>-4105</remarks>
		 [SupportByVersion("Excel", 12,14,15,16)]
		 xlChartElementPositionAutomatic = -4105,

		 /// <summary>
		 /// SupportByVersion Excel 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>-4114</remarks>
		 [SupportByVersion("Excel", 12,14,15,16)]
		 xlChartElementPositionCustom = -4114
	}
}