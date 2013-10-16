using System;
using NetOffice;
namespace NetOffice.ExcelApi.Enums
{
	 /// <summary>
	 /// SupportByVersion Excel 10, 11, 12, 14, 15
	 /// </summary>
	[SupportByVersionAttribute("Excel", 10,11,12,14,15)]
	[EntityTypeAttribute(EntityType.IsEnum)]
	public enum XlPrintErrors
	{
		 /// <summary>
		 /// SupportByVersion Excel 10, 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>0</remarks>
		 [SupportByVersionAttribute("Excel", 10,11,12,14,15)]
		 xlPrintErrorsDisplayed = 0,

		 /// <summary>
		 /// SupportByVersion Excel 10, 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>1</remarks>
		 [SupportByVersionAttribute("Excel", 10,11,12,14,15)]
		 xlPrintErrorsBlank = 1,

		 /// <summary>
		 /// SupportByVersion Excel 10, 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>2</remarks>
		 [SupportByVersionAttribute("Excel", 10,11,12,14,15)]
		 xlPrintErrorsDash = 2,

		 /// <summary>
		 /// SupportByVersion Excel 10, 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>3</remarks>
		 [SupportByVersionAttribute("Excel", 10,11,12,14,15)]
		 xlPrintErrorsNA = 3
	}
}