using System;
using NetOffice;
namespace NetOffice.WordApi.Enums
{
	 /// <summary>
	 /// SupportByVersion Word 10, 11, 12, 14, 15
	 /// </summary>
	 ///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff837297.aspx </remarks>
	[SupportByVersionAttribute("Word", 10,11,12,14,15)]
	[EntityTypeAttribute(EntityType.IsEnum)]
	public enum WdStyleSheetPrecedence
	{
		 /// <summary>
		 /// SupportByVersion Word 10, 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>-1</remarks>
		 [SupportByVersionAttribute("Word", 10,11,12,14,15)]
		 wdStyleSheetPrecedenceHigher = -1,

		 /// <summary>
		 /// SupportByVersion Word 10, 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>-2</remarks>
		 [SupportByVersionAttribute("Word", 10,11,12,14,15)]
		 wdStyleSheetPrecedenceLower = -2,

		 /// <summary>
		 /// SupportByVersion Word 10, 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>1</remarks>
		 [SupportByVersionAttribute("Word", 10,11,12,14,15)]
		 wdStyleSheetPrecedenceHighest = 1,

		 /// <summary>
		 /// SupportByVersion Word 10, 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>0</remarks>
		 [SupportByVersionAttribute("Word", 10,11,12,14,15)]
		 wdStyleSheetPrecedenceLowest = 0
	}
}