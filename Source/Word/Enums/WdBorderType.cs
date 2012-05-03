using System;
using NetOffice;
namespace NetOffice.WordApi.Enums
{
	 /// <summary>
	 /// SupportByVersion Word 9, 10, 11, 12, 14
	 /// </summary>
	[SupportByVersionAttribute("Word", 9,10,11,12,14)]
	[EntityTypeAttribute(EntityType.IsEnum)]
	public enum WdBorderType
	{
		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14
		 /// </summary>
		 /// <remarks>-1</remarks>
		 [SupportByVersionAttribute("Word", 9,10,11,12,14)]
		 wdBorderTop = -1,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14
		 /// </summary>
		 /// <remarks>-2</remarks>
		 [SupportByVersionAttribute("Word", 9,10,11,12,14)]
		 wdBorderLeft = -2,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14
		 /// </summary>
		 /// <remarks>-3</remarks>
		 [SupportByVersionAttribute("Word", 9,10,11,12,14)]
		 wdBorderBottom = -3,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14
		 /// </summary>
		 /// <remarks>-4</remarks>
		 [SupportByVersionAttribute("Word", 9,10,11,12,14)]
		 wdBorderRight = -4,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14
		 /// </summary>
		 /// <remarks>-5</remarks>
		 [SupportByVersionAttribute("Word", 9,10,11,12,14)]
		 wdBorderHorizontal = -5,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14
		 /// </summary>
		 /// <remarks>-6</remarks>
		 [SupportByVersionAttribute("Word", 9,10,11,12,14)]
		 wdBorderVertical = -6,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14
		 /// </summary>
		 /// <remarks>-7</remarks>
		 [SupportByVersionAttribute("Word", 9,10,11,12,14)]
		 wdBorderDiagonalDown = -7,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14
		 /// </summary>
		 /// <remarks>-8</remarks>
		 [SupportByVersionAttribute("Word", 9,10,11,12,14)]
		 wdBorderDiagonalUp = -8
	}
}