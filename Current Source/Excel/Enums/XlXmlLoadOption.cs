using System;
using LateBindingApi.Core;
namespace NetOffice.ExcelApi.Enums
{
	 /// <summary>
	 /// SupportByVersion Excel 11, 12, 14
	 /// </summary>
	[SupportByVersionAttribute("Excel", 11,12,14)]
	[EntityTypeAttribute(EntityType.IsEnum)]
	public enum XlXmlLoadOption
	{
		 /// <summary>
		 /// SupportByVersion Excel 11, 12, 14
		 /// </summary>
		 /// <remarks>0</remarks>
		 [SupportByVersionAttribute("Excel", 11,12,14)]
		 xlXmlLoadPromptUser = 0,

		 /// <summary>
		 /// SupportByVersion Excel 11, 12, 14
		 /// </summary>
		 /// <remarks>1</remarks>
		 [SupportByVersionAttribute("Excel", 11,12,14)]
		 xlXmlLoadOpenXml = 1,

		 /// <summary>
		 /// SupportByVersion Excel 11, 12, 14
		 /// </summary>
		 /// <remarks>2</remarks>
		 [SupportByVersionAttribute("Excel", 11,12,14)]
		 xlXmlLoadImportToList = 2,

		 /// <summary>
		 /// SupportByVersion Excel 11, 12, 14
		 /// </summary>
		 /// <remarks>3</remarks>
		 [SupportByVersionAttribute("Excel", 11,12,14)]
		 xlXmlLoadMapXml = 3
	}
}