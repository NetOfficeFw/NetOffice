﻿using System;
using NetOffice;
using NetOffice.Attributes;
namespace NetOffice.WordApi.Enums
{
	 /// <summary>
	 /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
	 /// </summary>
	 ///<remarks> MSDN Online Documentation: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.WdTablePosition"/> </remarks>
	[SupportByVersion("Word", 9,10,11,12,14,15,16)]
	[EntityType(EntityType.IsEnum)]
	public enum WdTablePosition
	{
		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>-999999</remarks>
		 [SupportByVersion("Word", 9,10,11,12,14,15,16)]
		 wdTableTop = -999999,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>-999998</remarks>
		 [SupportByVersion("Word", 9,10,11,12,14,15,16)]
		 wdTableLeft = -999998,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>-999997</remarks>
		 [SupportByVersion("Word", 9,10,11,12,14,15,16)]
		 wdTableBottom = -999997,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>-999996</remarks>
		 [SupportByVersion("Word", 9,10,11,12,14,15,16)]
		 wdTableRight = -999996,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>-999995</remarks>
		 [SupportByVersion("Word", 9,10,11,12,14,15,16)]
		 wdTableCenter = -999995,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>-999994</remarks>
		 [SupportByVersion("Word", 9,10,11,12,14,15,16)]
		 wdTableInside = -999994,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>-999993</remarks>
		 [SupportByVersion("Word", 9,10,11,12,14,15,16)]
		 wdTableOutside = -999993
	}
}