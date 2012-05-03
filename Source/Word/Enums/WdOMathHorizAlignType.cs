using System;
using NetOffice;
namespace NetOffice.WordApi.Enums
{
	 /// <summary>
	 /// SupportByVersion Word 12, 14
	 /// </summary>
	[SupportByVersionAttribute("Word", 12,14)]
	[EntityTypeAttribute(EntityType.IsEnum)]
	public enum WdOMathHorizAlignType
	{
		 /// <summary>
		 /// SupportByVersion Word 12, 14
		 /// </summary>
		 /// <remarks>0</remarks>
		 [SupportByVersionAttribute("Word", 12,14)]
		 wdOMathHorizAlignCenter = 0,

		 /// <summary>
		 /// SupportByVersion Word 12, 14
		 /// </summary>
		 /// <remarks>1</remarks>
		 [SupportByVersionAttribute("Word", 12,14)]
		 wdOMathHorizAlignLeft = 1,

		 /// <summary>
		 /// SupportByVersion Word 12, 14
		 /// </summary>
		 /// <remarks>2</remarks>
		 [SupportByVersionAttribute("Word", 12,14)]
		 wdOMathHorizAlignRight = 2
	}
}