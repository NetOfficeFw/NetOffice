using System;
using NetOffice;
using NetOffice.Attributes;
namespace NetOffice.ExcelApi.Enums
{
	 /// <summary>
	 /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
	 /// </summary>
	 ///<remarks> MSDN Online Documentation: <see href="https://docs.microsoft.com/en-us/office/vba/api/Excel.XlInsertShiftDirection"/> </remarks>
	[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
	[EntityType(EntityType.IsEnum)]
	public enum XlInsertShiftDirection
	{
		 /// <summary>
		 /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>-4121</remarks>
		 [SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		 xlShiftDown = -4121,

		 /// <summary>
		 /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>-4161</remarks>
		 [SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		 xlShiftToRight = -4161
	}
}