using System;
using NetOffice;
namespace NetOffice.ExcelApi.Enums
{
	 /// <summary>
	 /// SupportByVersion Excel 9, 10, 11, 12, 14
	 /// </summary>
	[SupportByVersionAttribute("Excel", 9,10,11,12,14)]
	[EntityTypeAttribute(EntityType.IsEnum)]
	public enum XlTextQualifier
	{
		 /// <summary>
		 /// SupportByVersion Excel 9, 10, 11, 12, 14
		 /// </summary>
		 /// <remarks>1</remarks>
		 [SupportByVersionAttribute("Excel", 9,10,11,12,14)]
		 xlTextQualifierDoubleQuote = 1,

		 /// <summary>
		 /// SupportByVersion Excel 9, 10, 11, 12, 14
		 /// </summary>
		 /// <remarks>-4142</remarks>
		 [SupportByVersionAttribute("Excel", 9,10,11,12,14)]
		 xlTextQualifierNone = -4142,

		 /// <summary>
		 /// SupportByVersion Excel 9, 10, 11, 12, 14
		 /// </summary>
		 /// <remarks>2</remarks>
		 [SupportByVersionAttribute("Excel", 9,10,11,12,14)]
		 xlTextQualifierSingleQuote = 2
	}
}