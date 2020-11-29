using System;
using NetOffice;
using NetOffice.Attributes;
namespace NetOffice.WordApi.Enums
{
	 /// <summary>
	 /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
	 /// </summary>
	 ///<remarks> MSDN Online Documentation: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.WdFramePosition"/> </remarks>
	[SupportByVersion("Word", 9,10,11,12,14,15,16)]
	[EntityType(EntityType.IsEnum)]
	public enum WdFramePosition
	{
		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>-999999</remarks>
		 [SupportByVersion("Word", 9,10,11,12,14,15,16)]
		 wdFrameTop = -999999,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>-999998</remarks>
		 [SupportByVersion("Word", 9,10,11,12,14,15,16)]
		 wdFrameLeft = -999998,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>-999997</remarks>
		 [SupportByVersion("Word", 9,10,11,12,14,15,16)]
		 wdFrameBottom = -999997,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>-999996</remarks>
		 [SupportByVersion("Word", 9,10,11,12,14,15,16)]
		 wdFrameRight = -999996,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>-999995</remarks>
		 [SupportByVersion("Word", 9,10,11,12,14,15,16)]
		 wdFrameCenter = -999995,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>-999994</remarks>
		 [SupportByVersion("Word", 9,10,11,12,14,15,16)]
		 wdFrameInside = -999994,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>-999993</remarks>
		 [SupportByVersion("Word", 9,10,11,12,14,15,16)]
		 wdFrameOutside = -999993
	}
}