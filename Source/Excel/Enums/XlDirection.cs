using System;
using NetOffice;
using NetOffice.Attributes;
namespace NetOffice.ExcelApi.Enums
{
	 /// <summary>
	 /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
	 /// </summary>
	 ///<remarks> MSDN Online Documentation: <see href="https://docs.microsoft.com/en-us/office/vba/api/Excel.XlDirection"/> </remarks>
	[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
	[EntityType(EntityType.IsEnum)]
	public enum XlDirection
	{
		 /// <summary>
		 /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>-4121</remarks>
		 [SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		 xlDown = -4121,

		 /// <summary>
		 /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>-4159</remarks>
		 [SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		 xlToLeft = -4159,

		 /// <summary>
		 /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>-4161</remarks>
		 [SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		 xlToRight = -4161,

		 /// <summary>
		 /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>-4162</remarks>
		 [SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		 xlUp = -4162
	}
}