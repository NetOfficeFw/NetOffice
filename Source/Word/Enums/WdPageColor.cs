using System;
using NetOffice;
using NetOffice.Attributes;
namespace NetOffice.WordApi.Enums
{
	 /// <summary>
	 /// SupportByVersion Word 15,16
	 /// </summary>
	 ///<remarks> MSDN Online Documentation: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.wdpagecolor"/> </remarks>
	[SupportByVersion("Word", 15, 16)]
	[EntityType(EntityType.IsEnum)]
	public enum WdPageColor
	{
		 /// <summary>
		 /// SupportByVersion Word 15,16
		 /// </summary>
		 /// <remarks>0</remarks>
		 [SupportByVersion("Word", 15, 16)]
		 wdPageColorNone = 0,

		 /// <summary>
		 /// SupportByVersion Word 15,16
		 /// </summary>
		 /// <remarks>1</remarks>
		 [SupportByVersion("Word", 15, 16)]
		 wdPageColorSepia = 1,

		 /// <summary>
		 /// SupportByVersion Word 15,16
		 /// </summary>
		 /// <remarks>2</remarks>
		 [SupportByVersion("Word", 15, 16)]
		 wdPageColorInverse = 2
	}
}