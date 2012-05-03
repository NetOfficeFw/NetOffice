using System;
using NetOffice;
namespace NetOffice.WordApi.Enums
{
	 /// <summary>
	 /// SupportByVersion Word 11, 12, 14
	 /// </summary>
	[SupportByVersionAttribute("Word", 11,12,14)]
	[EntityTypeAttribute(EntityType.IsEnum)]
	public enum WdXMLNodeLevel
	{
		 /// <summary>
		 /// SupportByVersion Word 11, 12, 14
		 /// </summary>
		 /// <remarks>0</remarks>
		 [SupportByVersionAttribute("Word", 11,12,14)]
		 wdXMLNodeLevelInline = 0,

		 /// <summary>
		 /// SupportByVersion Word 11, 12, 14
		 /// </summary>
		 /// <remarks>1</remarks>
		 [SupportByVersionAttribute("Word", 11,12,14)]
		 wdXMLNodeLevelParagraph = 1,

		 /// <summary>
		 /// SupportByVersion Word 11, 12, 14
		 /// </summary>
		 /// <remarks>2</remarks>
		 [SupportByVersionAttribute("Word", 11,12,14)]
		 wdXMLNodeLevelRow = 2,

		 /// <summary>
		 /// SupportByVersion Word 11, 12, 14
		 /// </summary>
		 /// <remarks>3</remarks>
		 [SupportByVersionAttribute("Word", 11,12,14)]
		 wdXMLNodeLevelCell = 3
	}
}