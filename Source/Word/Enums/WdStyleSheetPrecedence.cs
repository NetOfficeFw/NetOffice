using System;
using NetOffice;
namespace NetOffice.WordApi.Enums
{
	 /// <summary>
	 /// SupportByVersion Word 10, 11, 12, 14
	 /// </summary>
	[SupportByVersionAttribute("Word", 10,11,12,14)]
	[EntityTypeAttribute(EntityType.IsEnum)]
	public enum WdStyleSheetPrecedence
	{
		 /// <summary>
		 /// SupportByVersion Word 10, 11, 12, 14
		 /// </summary>
		 /// <remarks>-1</remarks>
		 [SupportByVersionAttribute("Word", 10,11,12,14)]
		 wdStyleSheetPrecedenceHigher = -1,

		 /// <summary>
		 /// SupportByVersion Word 10, 11, 12, 14
		 /// </summary>
		 /// <remarks>-2</remarks>
		 [SupportByVersionAttribute("Word", 10,11,12,14)]
		 wdStyleSheetPrecedenceLower = -2,

		 /// <summary>
		 /// SupportByVersion Word 10, 11, 12, 14
		 /// </summary>
		 /// <remarks>1</remarks>
		 [SupportByVersionAttribute("Word", 10,11,12,14)]
		 wdStyleSheetPrecedenceHighest = 1,

		 /// <summary>
		 /// SupportByVersion Word 10, 11, 12, 14
		 /// </summary>
		 /// <remarks>0</remarks>
		 [SupportByVersionAttribute("Word", 10,11,12,14)]
		 wdStyleSheetPrecedenceLowest = 0
	}
}