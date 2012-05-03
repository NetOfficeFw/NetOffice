using System;
using NetOffice;
namespace NetOffice.ExcelApi.Enums
{
	 /// <summary>
	 /// SupportByVersion Excel 9, 10, 11, 12, 14
	 /// </summary>
	[SupportByVersionAttribute("Excel", 9,10,11,12,14)]
	[EntityTypeAttribute(EntityType.IsEnum)]
	public enum XlCreator
	{
		 /// <summary>
		 /// SupportByVersion Excel 9, 10, 11, 12, 14
		 /// </summary>
		 /// <remarks>1480803660</remarks>
		 [SupportByVersionAttribute("Excel", 9,10,11,12,14)]
		 xlCreatorCode = 1480803660
	}
}