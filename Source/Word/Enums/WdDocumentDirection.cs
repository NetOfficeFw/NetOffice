using System;
using NetOffice;
namespace NetOffice.WordApi.Enums
{
	 /// <summary>
	 /// SupportByVersion Word 10, 11, 12, 14
	 /// </summary>
	[SupportByVersionAttribute("Word", 10,11,12,14)]
	[EntityTypeAttribute(EntityType.IsEnum)]
	public enum WdDocumentDirection
	{
		 /// <summary>
		 /// SupportByVersion Word 10, 11, 12, 14
		 /// </summary>
		 /// <remarks>0</remarks>
		 [SupportByVersionAttribute("Word", 10,11,12,14)]
		 wdLeftToRight = 0,

		 /// <summary>
		 /// SupportByVersion Word 10, 11, 12, 14
		 /// </summary>
		 /// <remarks>1</remarks>
		 [SupportByVersionAttribute("Word", 10,11,12,14)]
		 wdRightToLeft = 1
	}
}