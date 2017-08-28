using System;
using NetOffice;
using NetOffice.Attributes;
namespace NetOffice.WordApi.Enums
{
	 /// <summary>
	 /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
	 /// </summary>
	[SupportByVersion("Word", 9,10,11,12,14,15,16)]
	[EntityType(EntityType.IsEnum)]
	public enum WdSummaryLength
	{
		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>-2</remarks>
		 [SupportByVersion("Word", 9,10,11,12,14,15,16)]
		 wd10Sentences = -2,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>-3</remarks>
		 [SupportByVersion("Word", 9,10,11,12,14,15,16)]
		 wd20Sentences = -3,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>-4</remarks>
		 [SupportByVersion("Word", 9,10,11,12,14,15,16)]
		 wd100Words = -4,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>-5</remarks>
		 [SupportByVersion("Word", 9,10,11,12,14,15,16)]
		 wd500Words = -5,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>-6</remarks>
		 [SupportByVersion("Word", 9,10,11,12,14,15,16)]
		 wd10Percent = -6,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>-7</remarks>
		 [SupportByVersion("Word", 9,10,11,12,14,15,16)]
		 wd25Percent = -7,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>-8</remarks>
		 [SupportByVersion("Word", 9,10,11,12,14,15,16)]
		 wd50Percent = -8,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>-9</remarks>
		 [SupportByVersion("Word", 9,10,11,12,14,15,16)]
		 wd75Percent = -9
	}
}