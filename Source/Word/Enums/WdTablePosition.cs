using System;
using NetOffice;
namespace NetOffice.WordApi.Enums
{
	 /// <summary>
	 /// SupportByVersion Word 9, 10, 11, 12, 14, 15
	 /// </summary>
	 ///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff835469.aspx </remarks>
	[SupportByVersionAttribute("Word", 9,10,11,12,14,15)]
	[EntityTypeAttribute(EntityType.IsEnum)]
	public enum WdTablePosition
	{
		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>-999999</remarks>
		 [SupportByVersionAttribute("Word", 9,10,11,12,14,15)]
		 wdTableTop = -999999,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>-999998</remarks>
		 [SupportByVersionAttribute("Word", 9,10,11,12,14,15)]
		 wdTableLeft = -999998,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>-999997</remarks>
		 [SupportByVersionAttribute("Word", 9,10,11,12,14,15)]
		 wdTableBottom = -999997,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>-999996</remarks>
		 [SupportByVersionAttribute("Word", 9,10,11,12,14,15)]
		 wdTableRight = -999996,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>-999995</remarks>
		 [SupportByVersionAttribute("Word", 9,10,11,12,14,15)]
		 wdTableCenter = -999995,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>-999994</remarks>
		 [SupportByVersionAttribute("Word", 9,10,11,12,14,15)]
		 wdTableInside = -999994,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>-999993</remarks>
		 [SupportByVersionAttribute("Word", 9,10,11,12,14,15)]
		 wdTableOutside = -999993
	}
}