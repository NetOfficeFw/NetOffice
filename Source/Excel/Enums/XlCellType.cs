using System;
using NetOffice;
using NetOffice.Attributes;
namespace NetOffice.ExcelApi.Enums
{
	 /// <summary>
	 /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
	 /// </summary>
	 ///<remarks> MSDN Online Documentation: <see href="https://docs.microsoft.com/en-us/office/vba/api/Excel.XlCellType"/> </remarks>
	[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
	[EntityType(EntityType.IsEnum)]
	public enum XlCellType
	{
		 /// <summary>
		 /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>4</remarks>
		 [SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		 xlCellTypeBlanks = 4,

		 /// <summary>
		 /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>2</remarks>
		 [SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		 xlCellTypeConstants = 2,

		 /// <summary>
		 /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>-4123</remarks>
		 [SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		 xlCellTypeFormulas = -4123,

		 /// <summary>
		 /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>11</remarks>
		 [SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		 xlCellTypeLastCell = 11,

		 /// <summary>
		 /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>-4144</remarks>
		 [SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		 xlCellTypeComments = -4144,

		 /// <summary>
		 /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>12</remarks>
		 [SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		 xlCellTypeVisible = 12,

		 /// <summary>
		 /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>-4172</remarks>
		 [SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		 xlCellTypeAllFormatConditions = -4172,

		 /// <summary>
		 /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>-4173</remarks>
		 [SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		 xlCellTypeSameFormatConditions = -4173,

		 /// <summary>
		 /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>-4174</remarks>
		 [SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		 xlCellTypeAllValidation = -4174,

		 /// <summary>
		 /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>-4175</remarks>
		 [SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		 xlCellTypeSameValidation = -4175
	}
}