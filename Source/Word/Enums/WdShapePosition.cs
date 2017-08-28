using System;
using NetOffice;
using NetOffice.Attributes;
namespace NetOffice.WordApi.Enums
{
	 /// <summary>
	 /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
	 /// </summary>
	 ///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff197284.aspx </remarks>
	[SupportByVersion("Word", 9,10,11,12,14,15,16)]
	[EntityType(EntityType.IsEnum)]
	public enum WdShapePosition
	{
		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>-999999</remarks>
		 [SupportByVersion("Word", 9,10,11,12,14,15,16)]
		 wdShapeTop = -999999,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>-999998</remarks>
		 [SupportByVersion("Word", 9,10,11,12,14,15,16)]
		 wdShapeLeft = -999998,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>-999997</remarks>
		 [SupportByVersion("Word", 9,10,11,12,14,15,16)]
		 wdShapeBottom = -999997,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>-999996</remarks>
		 [SupportByVersion("Word", 9,10,11,12,14,15,16)]
		 wdShapeRight = -999996,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>-999995</remarks>
		 [SupportByVersion("Word", 9,10,11,12,14,15,16)]
		 wdShapeCenter = -999995,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>-999994</remarks>
		 [SupportByVersion("Word", 9,10,11,12,14,15,16)]
		 wdShapeInside = -999994,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>-999993</remarks>
		 [SupportByVersion("Word", 9,10,11,12,14,15,16)]
		 wdShapeOutside = -999993
	}
}