using System;
using NetOffice;
using NetOffice.Attributes;
namespace NetOffice.ExcelApi.Enums
{
	 /// <summary>
	 /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
	 /// </summary>
	 ///<remarks> MSDN Online Documentation: <see href="https://docs.microsoft.com/en-us/office/vba/api/Excel.XlVAlign"/> </remarks>
	[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
	[EntityType(EntityType.IsEnum)]
	public enum XlVAlign
	{
		 /// <summary>
		 /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>-4107</remarks>
		 [SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		 xlVAlignBottom = -4107,

		 /// <summary>
		 /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>-4108</remarks>
		 [SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		 xlVAlignCenter = -4108,

		 /// <summary>
		 /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>-4117</remarks>
		 [SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		 xlVAlignDistributed = -4117,

		 /// <summary>
		 /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>-4130</remarks>
		 [SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		 xlVAlignJustify = -4130,

		 /// <summary>
		 /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>-4160</remarks>
		 [SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		 xlVAlignTop = -4160
	}
}