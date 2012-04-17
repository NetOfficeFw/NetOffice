using System;
using LateBindingApi.Core;
namespace NetOffice.ExcelApi.Enums
{
	 /// <summary>
	 /// SupportByVersion Excel 12, 14
	 /// </summary>
	[SupportByVersionAttribute("Excel", 12,14)]
	[EntityTypeAttribute(EntityType.IsEnum)]
	public enum XlChartElementPosition
	{
		 /// <summary>
		 /// SupportByVersion Excel 12, 14
		 /// </summary>
		 /// <remarks>-4105</remarks>
		 [SupportByVersionAttribute("Excel", 12,14)]
		 xlChartElementPositionAutomatic = -4105,

		 /// <summary>
		 /// SupportByVersion Excel 12, 14
		 /// </summary>
		 /// <remarks>-4114</remarks>
		 [SupportByVersionAttribute("Excel", 12,14)]
		 xlChartElementPositionCustom = -4114
	}
}