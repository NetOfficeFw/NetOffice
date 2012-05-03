using System;
using NetOffice;
namespace NetOffice.WordApi.Enums
{
	 /// <summary>
	 /// SupportByVersion Word 14
	 /// </summary>
	[SupportByVersionAttribute("Word", 14)]
	[EntityTypeAttribute(EntityType.IsEnum)]
	public enum WdCompatibilityMode
	{
		 /// <summary>
		 /// SupportByVersion Word 14
		 /// </summary>
		 /// <remarks>11</remarks>
		 [SupportByVersionAttribute("Word", 14)]
		 wdWord2003 = 11,

		 /// <summary>
		 /// SupportByVersion Word 14
		 /// </summary>
		 /// <remarks>12</remarks>
		 [SupportByVersionAttribute("Word", 14)]
		 wdWord2007 = 12,

		 /// <summary>
		 /// SupportByVersion Word 14
		 /// </summary>
		 /// <remarks>14</remarks>
		 [SupportByVersionAttribute("Word", 14)]
		 wdWord2010 = 14,

		 /// <summary>
		 /// SupportByVersion Word 14
		 /// </summary>
		 /// <remarks>65535</remarks>
		 [SupportByVersionAttribute("Word", 14)]
		 wdCurrent = 65535
	}
}