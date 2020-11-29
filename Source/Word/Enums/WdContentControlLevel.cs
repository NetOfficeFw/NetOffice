using System;
using NetOffice;
using NetOffice.Attributes;
namespace NetOffice.WordApi.Enums
{
	 /// <summary>
	 /// SupportByVersion Word 15,16
	 /// </summary>
	 ///<remarks> MSDN Online Documentation: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.wdcontentcontrollevel"/> </remarks>
	[SupportByVersion("Word", 15, 16)]
	[EntityType(EntityType.IsEnum)]
	public enum WdContentControlLevel
	{
		 /// <summary>
		 /// SupportByVersion Word 15,16
		 /// </summary>
		 /// <remarks>0</remarks>
		 [SupportByVersion("Word", 15, 16)]
		 wdContentControlLevelInline = 0,

		 /// <summary>
		 /// SupportByVersion Word 15,16
		 /// </summary>
		 /// <remarks>1</remarks>
		 [SupportByVersion("Word", 15, 16)]
		 wdContentControlLevelParagraph = 1,

		 /// <summary>
		 /// SupportByVersion Word 15,16
		 /// </summary>
		 /// <remarks>2</remarks>
		 [SupportByVersion("Word", 15, 16)]
		 wdContentControlLevelRow = 2,

		 /// <summary>
		 /// SupportByVersion Word 15,16
		 /// </summary>
		 /// <remarks>3</remarks>
		 [SupportByVersion("Word", 15, 16)]
		 wdContentControlLevelCell = 3
	}
}