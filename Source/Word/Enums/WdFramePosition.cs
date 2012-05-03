using System;
using NetOffice;
namespace NetOffice.WordApi.Enums
{
	 /// <summary>
	 /// SupportByVersion Word 9, 10, 11, 12, 14
	 /// </summary>
	[SupportByVersionAttribute("Word", 9,10,11,12,14)]
	[EntityTypeAttribute(EntityType.IsEnum)]
	public enum WdFramePosition
	{
		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14
		 /// </summary>
		 /// <remarks>-999999</remarks>
		 [SupportByVersionAttribute("Word", 9,10,11,12,14)]
		 wdFrameTop = -999999,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14
		 /// </summary>
		 /// <remarks>-999998</remarks>
		 [SupportByVersionAttribute("Word", 9,10,11,12,14)]
		 wdFrameLeft = -999998,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14
		 /// </summary>
		 /// <remarks>-999997</remarks>
		 [SupportByVersionAttribute("Word", 9,10,11,12,14)]
		 wdFrameBottom = -999997,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14
		 /// </summary>
		 /// <remarks>-999996</remarks>
		 [SupportByVersionAttribute("Word", 9,10,11,12,14)]
		 wdFrameRight = -999996,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14
		 /// </summary>
		 /// <remarks>-999995</remarks>
		 [SupportByVersionAttribute("Word", 9,10,11,12,14)]
		 wdFrameCenter = -999995,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14
		 /// </summary>
		 /// <remarks>-999994</remarks>
		 [SupportByVersionAttribute("Word", 9,10,11,12,14)]
		 wdFrameInside = -999994,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14
		 /// </summary>
		 /// <remarks>-999993</remarks>
		 [SupportByVersionAttribute("Word", 9,10,11,12,14)]
		 wdFrameOutside = -999993
	}
}