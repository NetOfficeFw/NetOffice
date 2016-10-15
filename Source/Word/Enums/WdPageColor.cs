using System;
using NetOffice;
namespace NetOffice.WordApi.Enums
{
	 /// <summary>
	 /// SupportByVersion Word 15,16
	 /// </summary>
	 ///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/jj231170.aspx </remarks>
	[SupportByVersionAttribute("Word", 15, 16)]
	[EntityTypeAttribute(EntityType.IsEnum)]
	public enum WdPageColor
	{
		 /// <summary>
		 /// SupportByVersion Word 15,16
		 /// </summary>
		 /// <remarks>0</remarks>
		 [SupportByVersionAttribute("Word", 15, 16)]
		 wdPageColorNone = 0,

		 /// <summary>
		 /// SupportByVersion Word 15,16
		 /// </summary>
		 /// <remarks>1</remarks>
		 [SupportByVersionAttribute("Word", 15, 16)]
		 wdPageColorSepia = 1,

		 /// <summary>
		 /// SupportByVersion Word 15,16
		 /// </summary>
		 /// <remarks>2</remarks>
		 [SupportByVersionAttribute("Word", 15, 16)]
		 wdPageColorInverse = 2
	}
}