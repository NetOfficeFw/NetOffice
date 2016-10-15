using System;
using NetOffice;
namespace NetOffice.WordApi.Enums
{
	 /// <summary>
	 /// SupportByVersion Word 14, 15, 16
	 /// </summary>
	 ///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff192388.aspx </remarks>
	[SupportByVersionAttribute("Word", 14,15,16)]
	[EntityTypeAttribute(EntityType.IsEnum)]
	public enum WdCompatibilityMode
	{
		 /// <summary>
		 /// SupportByVersion Word 14, 15, 16
		 /// </summary>
		 /// <remarks>11</remarks>
		 [SupportByVersionAttribute("Word", 14,15,16)]
		 wdWord2003 = 11,

		 /// <summary>
		 /// SupportByVersion Word 14, 15, 16
		 /// </summary>
		 /// <remarks>12</remarks>
		 [SupportByVersionAttribute("Word", 14,15,16)]
		 wdWord2007 = 12,

		 /// <summary>
		 /// SupportByVersion Word 14, 15, 16
		 /// </summary>
		 /// <remarks>14</remarks>
		 [SupportByVersionAttribute("Word", 14,15,16)]
		 wdWord2010 = 14,

		 /// <summary>
		 /// SupportByVersion Word 14, 15, 16
		 /// </summary>
		 /// <remarks>65535</remarks>
		 [SupportByVersionAttribute("Word", 14,15,16)]
		 wdCurrent = 65535,

		 /// <summary>
		 /// SupportByVersion Word 15,16
		 /// </summary>
		 /// <remarks>15</remarks>
		 [SupportByVersionAttribute("Word", 15, 16)]
		 wdWord2012 = 15
	}
}