using System;
using NetOffice;
using NetOffice.Attributes;
namespace NetOffice.WordApi.Enums
{
	 /// <summary>
	 /// SupportByVersion Word 14, 15, 16
	 /// </summary>
	 ///<remarks> MSDN Online Documentation: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.WdCompatibilityMode"/> </remarks>
	[SupportByVersion("Word", 14,15,16)]
	[EntityType(EntityType.IsEnum)]
	public enum WdCompatibilityMode
	{
		 /// <summary>
		 /// SupportByVersion Word 14, 15, 16
		 /// </summary>
		 /// <remarks>11</remarks>
		 [SupportByVersion("Word", 14,15,16)]
		 wdWord2003 = 11,

		 /// <summary>
		 /// SupportByVersion Word 14, 15, 16
		 /// </summary>
		 /// <remarks>12</remarks>
		 [SupportByVersion("Word", 14,15,16)]
		 wdWord2007 = 12,

		 /// <summary>
		 /// SupportByVersion Word 14, 15, 16
		 /// </summary>
		 /// <remarks>14</remarks>
		 [SupportByVersion("Word", 14,15,16)]
		 wdWord2010 = 14,

		 /// <summary>
		 /// SupportByVersion Word 14, 15, 16
		 /// </summary>
		 /// <remarks>65535</remarks>
		 [SupportByVersion("Word", 14,15,16)]
		 wdCurrent = 65535,

		 /// <summary>
		 /// SupportByVersion Word 15,16
		 /// </summary>
		 /// <remarks>15</remarks>
		 [SupportByVersion("Word", 15, 16)]
		 wdWord2012 = 15
	}
}