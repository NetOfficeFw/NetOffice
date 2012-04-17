using System;
using LateBindingApi.Core;
namespace NetOffice.ExcelApi.Enums
{
	 /// <summary>
	 /// SupportByVersion Excel 9, 10, 11, 12, 14
	 /// </summary>
	[SupportByVersionAttribute("Excel", 9,10,11,12,14)]
	[EntityTypeAttribute(EntityType.IsEnum)]
	public enum XlSummaryColumn
	{
		 /// <summary>
		 /// SupportByVersion Excel 9, 10, 11, 12, 14
		 /// </summary>
		 /// <remarks>-4131</remarks>
		 [SupportByVersionAttribute("Excel", 9,10,11,12,14)]
		 xlSummaryOnLeft = -4131,

		 /// <summary>
		 /// SupportByVersion Excel 9, 10, 11, 12, 14
		 /// </summary>
		 /// <remarks>-4152</remarks>
		 [SupportByVersionAttribute("Excel", 9,10,11,12,14)]
		 xlSummaryOnRight = -4152
	}
}