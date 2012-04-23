using System;
using NetOffice;
namespace NetOffice.ExcelApi.Enums
{
	 /// <summary>
	 /// SupportByVersion Excel 9, 10, 11, 12, 14
	 /// </summary>
	[SupportByVersionAttribute("Excel", 9,10,11,12,14)]
	[EntityTypeAttribute(EntityType.IsEnum)]
	public enum XlOrientation
	{
		 /// <summary>
		 /// SupportByVersion Excel 9, 10, 11, 12, 14
		 /// </summary>
		 /// <remarks>-4170</remarks>
		 [SupportByVersionAttribute("Excel", 9,10,11,12,14)]
		 xlDownward = -4170,

		 /// <summary>
		 /// SupportByVersion Excel 9, 10, 11, 12, 14
		 /// </summary>
		 /// <remarks>-4128</remarks>
		 [SupportByVersionAttribute("Excel", 9,10,11,12,14)]
		 xlHorizontal = -4128,

		 /// <summary>
		 /// SupportByVersion Excel 9, 10, 11, 12, 14
		 /// </summary>
		 /// <remarks>-4171</remarks>
		 [SupportByVersionAttribute("Excel", 9,10,11,12,14)]
		 xlUpward = -4171,

		 /// <summary>
		 /// SupportByVersion Excel 9, 10, 11, 12, 14
		 /// </summary>
		 /// <remarks>-4166</remarks>
		 [SupportByVersionAttribute("Excel", 9,10,11,12,14)]
		 xlVertical = -4166
	}
}