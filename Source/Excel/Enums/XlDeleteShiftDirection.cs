using System;
using NetOffice;
using NetOffice.Attributes;
namespace NetOffice.ExcelApi.Enums
{
	 /// <summary>
	 /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
	 /// </summary>
	 ///<remarks> MSDN Online Documentation: <see href="https://docs.microsoft.com/en-us/office/vba/api/Excel.XlDeleteShiftDirection"/> </remarks>
	[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
	[EntityType(EntityType.IsEnum)]
	public enum XlDeleteShiftDirection
	{
		 /// <summary>
		 /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>-4159</remarks>
		 [SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		 xlShiftToLeft = -4159,

		 /// <summary>
		 /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>-4162</remarks>
		 [SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		 xlShiftUp = -4162
	}
}