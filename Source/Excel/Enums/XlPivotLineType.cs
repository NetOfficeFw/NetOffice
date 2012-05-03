using System;
using NetOffice;
namespace NetOffice.ExcelApi.Enums
{
	 /// <summary>
	 /// SupportByVersion Excel 12, 14
	 /// </summary>
	[SupportByVersionAttribute("Excel", 12,14)]
	[EntityTypeAttribute(EntityType.IsEnum)]
	public enum XlPivotLineType
	{
		 /// <summary>
		 /// SupportByVersion Excel 12, 14
		 /// </summary>
		 /// <remarks>0</remarks>
		 [SupportByVersionAttribute("Excel", 12,14)]
		 xlPivotLineRegular = 0,

		 /// <summary>
		 /// SupportByVersion Excel 12, 14
		 /// </summary>
		 /// <remarks>1</remarks>
		 [SupportByVersionAttribute("Excel", 12,14)]
		 xlPivotLineSubtotal = 1,

		 /// <summary>
		 /// SupportByVersion Excel 12, 14
		 /// </summary>
		 /// <remarks>2</remarks>
		 [SupportByVersionAttribute("Excel", 12,14)]
		 xlPivotLineGrandTotal = 2,

		 /// <summary>
		 /// SupportByVersion Excel 12, 14
		 /// </summary>
		 /// <remarks>3</remarks>
		 [SupportByVersionAttribute("Excel", 12,14)]
		 xlPivotLineBlank = 3
	}
}