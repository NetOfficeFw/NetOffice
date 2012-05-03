using System;
using NetOffice;
namespace NetOffice.ExcelApi.Enums
{
	 /// <summary>
	 /// SupportByVersion Excel 9, 10, 11, 12, 14
	 /// </summary>
	[SupportByVersionAttribute("Excel", 9,10,11,12,14)]
	[EntityTypeAttribute(EntityType.IsEnum)]
	public enum XlErrorBarType
	{
		 /// <summary>
		 /// SupportByVersion Excel 9, 10, 11, 12, 14
		 /// </summary>
		 /// <remarks>-4114</remarks>
		 [SupportByVersionAttribute("Excel", 9,10,11,12,14)]
		 xlErrorBarTypeCustom = -4114,

		 /// <summary>
		 /// SupportByVersion Excel 9, 10, 11, 12, 14
		 /// </summary>
		 /// <remarks>1</remarks>
		 [SupportByVersionAttribute("Excel", 9,10,11,12,14)]
		 xlErrorBarTypeFixedValue = 1,

		 /// <summary>
		 /// SupportByVersion Excel 9, 10, 11, 12, 14
		 /// </summary>
		 /// <remarks>2</remarks>
		 [SupportByVersionAttribute("Excel", 9,10,11,12,14)]
		 xlErrorBarTypePercent = 2,

		 /// <summary>
		 /// SupportByVersion Excel 9, 10, 11, 12, 14
		 /// </summary>
		 /// <remarks>-4155</remarks>
		 [SupportByVersionAttribute("Excel", 9,10,11,12,14)]
		 xlErrorBarTypeStDev = -4155,

		 /// <summary>
		 /// SupportByVersion Excel 9, 10, 11, 12, 14
		 /// </summary>
		 /// <remarks>4</remarks>
		 [SupportByVersionAttribute("Excel", 9,10,11,12,14)]
		 xlErrorBarTypeStError = 4
	}
}