using System;
using NetOffice;
namespace NetOffice.WordApi.Enums
{
	 /// <summary>
	 /// SupportByVersion Word 15
	 /// </summary>
	 ///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/jj231523.aspx </remarks>
	[SupportByVersionAttribute("Word", 15)]
	[EntityTypeAttribute(EntityType.IsEnum)]
	public enum WdContentControlLevel
	{
		 /// <summary>
		 /// SupportByVersion Word 15
		 /// </summary>
		 /// <remarks>0</remarks>
		 [SupportByVersionAttribute("Word", 15)]
		 wdContentControlLevelInline = 0,

		 /// <summary>
		 /// SupportByVersion Word 15
		 /// </summary>
		 /// <remarks>1</remarks>
		 [SupportByVersionAttribute("Word", 15)]
		 wdContentControlLevelParagraph = 1,

		 /// <summary>
		 /// SupportByVersion Word 15
		 /// </summary>
		 /// <remarks>2</remarks>
		 [SupportByVersionAttribute("Word", 15)]
		 wdContentControlLevelRow = 2,

		 /// <summary>
		 /// SupportByVersion Word 15
		 /// </summary>
		 /// <remarks>3</remarks>
		 [SupportByVersionAttribute("Word", 15)]
		 wdContentControlLevelCell = 3
	}
}