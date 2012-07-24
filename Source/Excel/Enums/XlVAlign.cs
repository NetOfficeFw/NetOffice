using System;
using NetOffice;
namespace NetOffice.ExcelApi.Enums
{
	 /// <summary>
	 /// SupportByVersion Excel 9, 10, 11, 12, 14, 15
	 /// </summary>
	[SupportByVersionAttribute("Excel", 9,10,11,12,14,15)]
	[EntityTypeAttribute(EntityType.IsEnum)]
	public enum XlVAlign
	{
		 /// <summary>
		 /// SupportByVersion Excel 9, 10, 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>-4107</remarks>
		 [SupportByVersionAttribute("Excel", 9,10,11,12,14,15)]
		 xlVAlignBottom = -4107,

		 /// <summary>
		 /// SupportByVersion Excel 9, 10, 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>-4108</remarks>
		 [SupportByVersionAttribute("Excel", 9,10,11,12,14,15)]
		 xlVAlignCenter = -4108,

		 /// <summary>
		 /// SupportByVersion Excel 9, 10, 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>-4117</remarks>
		 [SupportByVersionAttribute("Excel", 9,10,11,12,14,15)]
		 xlVAlignDistributed = -4117,

		 /// <summary>
		 /// SupportByVersion Excel 9, 10, 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>-4130</remarks>
		 [SupportByVersionAttribute("Excel", 9,10,11,12,14,15)]
		 xlVAlignJustify = -4130,

		 /// <summary>
		 /// SupportByVersion Excel 9, 10, 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>-4160</remarks>
		 [SupportByVersionAttribute("Excel", 9,10,11,12,14,15)]
		 xlVAlignTop = -4160
	}
}