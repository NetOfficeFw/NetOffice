using System;
using NetOffice;
using NetOffice.Attributes;
namespace NetOffice.WordApi.Enums
{
	 /// <summary>
	 /// SupportByVersion Word 11, 12, 14, 15, 16
	 /// </summary>
	[SupportByVersion("Word", 11,12,14,15,16)]
	[EntityType(EntityType.IsEnum)]
	public enum WdXMLNodeLevel
	{
		 /// <summary>
		 /// SupportByVersion Word 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>0</remarks>
		 [SupportByVersion("Word", 11,12,14,15,16)]
		 wdXMLNodeLevelInline = 0,

		 /// <summary>
		 /// SupportByVersion Word 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>1</remarks>
		 [SupportByVersion("Word", 11,12,14,15,16)]
		 wdXMLNodeLevelParagraph = 1,

		 /// <summary>
		 /// SupportByVersion Word 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>2</remarks>
		 [SupportByVersion("Word", 11,12,14,15,16)]
		 wdXMLNodeLevelRow = 2,

		 /// <summary>
		 /// SupportByVersion Word 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>3</remarks>
		 [SupportByVersion("Word", 11,12,14,15,16)]
		 wdXMLNodeLevelCell = 3
	}
}