using System;
using NetOffice;
namespace NetOffice.ExcelApi.Enums
{
	 /// <summary>
	 /// SupportByVersion Excel 10, 11, 12, 14
	 /// </summary>
	[SupportByVersionAttribute("Excel", 10,11,12,14)]
	[EntityTypeAttribute(EntityType.IsEnum)]
	public enum XlCorruptLoad
	{
		 /// <summary>
		 /// SupportByVersion Excel 10, 11, 12, 14
		 /// </summary>
		 /// <remarks>0</remarks>
		 [SupportByVersionAttribute("Excel", 10,11,12,14)]
		 xlNormalLoad = 0,

		 /// <summary>
		 /// SupportByVersion Excel 10, 11, 12, 14
		 /// </summary>
		 /// <remarks>1</remarks>
		 [SupportByVersionAttribute("Excel", 10,11,12,14)]
		 xlRepairFile = 1,

		 /// <summary>
		 /// SupportByVersion Excel 10, 11, 12, 14
		 /// </summary>
		 /// <remarks>2</remarks>
		 [SupportByVersionAttribute("Excel", 10,11,12,14)]
		 xlExtractData = 2
	}
}