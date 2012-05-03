using System;
using NetOffice;
namespace NetOffice.ExcelApi.Enums
{
	 /// <summary>
	 /// SupportByVersion Excel 10, 11, 12, 14
	 /// </summary>
	[SupportByVersionAttribute("Excel", 10,11,12,14)]
	[EntityTypeAttribute(EntityType.IsEnum)]
	public enum XlRangeValueDataType
	{
		 /// <summary>
		 /// SupportByVersion Excel 10, 11, 12, 14
		 /// </summary>
		 /// <remarks>10</remarks>
		 [SupportByVersionAttribute("Excel", 10,11,12,14)]
		 xlRangeValueDefault = 10,

		 /// <summary>
		 /// SupportByVersion Excel 10, 11, 12, 14
		 /// </summary>
		 /// <remarks>11</remarks>
		 [SupportByVersionAttribute("Excel", 10,11,12,14)]
		 xlRangeValueXMLSpreadsheet = 11,

		 /// <summary>
		 /// SupportByVersion Excel 10, 11, 12, 14
		 /// </summary>
		 /// <remarks>12</remarks>
		 [SupportByVersionAttribute("Excel", 10,11,12,14)]
		 xlRangeValueMSPersistXML = 12
	}
}