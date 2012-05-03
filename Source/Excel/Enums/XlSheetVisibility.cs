using System;
using NetOffice;
namespace NetOffice.ExcelApi.Enums
{
	 /// <summary>
	 /// SupportByVersion Excel 9, 10, 11, 12, 14
	 /// </summary>
	[SupportByVersionAttribute("Excel", 9,10,11,12,14)]
	[EntityTypeAttribute(EntityType.IsEnum)]
	public enum XlSheetVisibility
	{
		 /// <summary>
		 /// SupportByVersion Excel 9, 10, 11, 12, 14
		 /// </summary>
		 /// <remarks>-1</remarks>
		 [SupportByVersionAttribute("Excel", 9,10,11,12,14)]
		 xlSheetVisible = -1,

		 /// <summary>
		 /// SupportByVersion Excel 9, 10, 11, 12, 14
		 /// </summary>
		 /// <remarks>0</remarks>
		 [SupportByVersionAttribute("Excel", 9,10,11,12,14)]
		 xlSheetHidden = 0,

		 /// <summary>
		 /// SupportByVersion Excel 9, 10, 11, 12, 14
		 /// </summary>
		 /// <remarks>2</remarks>
		 [SupportByVersionAttribute("Excel", 9,10,11,12,14)]
		 xlSheetVeryHidden = 2
	}
}