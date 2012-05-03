using System;
using NetOffice;
namespace NetOffice.ExcelApi.Enums
{
	 /// <summary>
	 /// SupportByVersion Excel 9, 10, 11, 12, 14
	 /// </summary>
	[SupportByVersionAttribute("Excel", 9,10,11,12,14)]
	[EntityTypeAttribute(EntityType.IsEnum)]
	public enum XlCellType
	{
		 /// <summary>
		 /// SupportByVersion Excel 9, 10, 11, 12, 14
		 /// </summary>
		 /// <remarks>4</remarks>
		 [SupportByVersionAttribute("Excel", 9,10,11,12,14)]
		 xlCellTypeBlanks = 4,

		 /// <summary>
		 /// SupportByVersion Excel 9, 10, 11, 12, 14
		 /// </summary>
		 /// <remarks>2</remarks>
		 [SupportByVersionAttribute("Excel", 9,10,11,12,14)]
		 xlCellTypeConstants = 2,

		 /// <summary>
		 /// SupportByVersion Excel 9, 10, 11, 12, 14
		 /// </summary>
		 /// <remarks>-4123</remarks>
		 [SupportByVersionAttribute("Excel", 9,10,11,12,14)]
		 xlCellTypeFormulas = -4123,

		 /// <summary>
		 /// SupportByVersion Excel 9, 10, 11, 12, 14
		 /// </summary>
		 /// <remarks>11</remarks>
		 [SupportByVersionAttribute("Excel", 9,10,11,12,14)]
		 xlCellTypeLastCell = 11,

		 /// <summary>
		 /// SupportByVersion Excel 9, 10, 11, 12, 14
		 /// </summary>
		 /// <remarks>-4144</remarks>
		 [SupportByVersionAttribute("Excel", 9,10,11,12,14)]
		 xlCellTypeComments = -4144,

		 /// <summary>
		 /// SupportByVersion Excel 9, 10, 11, 12, 14
		 /// </summary>
		 /// <remarks>12</remarks>
		 [SupportByVersionAttribute("Excel", 9,10,11,12,14)]
		 xlCellTypeVisible = 12,

		 /// <summary>
		 /// SupportByVersion Excel 9, 10, 11, 12, 14
		 /// </summary>
		 /// <remarks>-4172</remarks>
		 [SupportByVersionAttribute("Excel", 9,10,11,12,14)]
		 xlCellTypeAllFormatConditions = -4172,

		 /// <summary>
		 /// SupportByVersion Excel 9, 10, 11, 12, 14
		 /// </summary>
		 /// <remarks>-4173</remarks>
		 [SupportByVersionAttribute("Excel", 9,10,11,12,14)]
		 xlCellTypeSameFormatConditions = -4173,

		 /// <summary>
		 /// SupportByVersion Excel 9, 10, 11, 12, 14
		 /// </summary>
		 /// <remarks>-4174</remarks>
		 [SupportByVersionAttribute("Excel", 9,10,11,12,14)]
		 xlCellTypeAllValidation = -4174,

		 /// <summary>
		 /// SupportByVersion Excel 9, 10, 11, 12, 14
		 /// </summary>
		 /// <remarks>-4175</remarks>
		 [SupportByVersionAttribute("Excel", 9,10,11,12,14)]
		 xlCellTypeSameValidation = -4175
	}
}