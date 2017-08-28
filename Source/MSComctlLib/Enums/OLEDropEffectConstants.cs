using System;
using NetOffice;
using NetOffice.Attributes;
namespace NetOffice.MSComctlLibApi.Enums
{
	 /// <summary>
	 /// SupportByVersion MSComctlLib 6
	 /// </summary>
	[SupportByVersion("MSComctlLib", 6)]
	[EntityType(EntityType.IsEnum)]
	public enum OLEDropEffectConstants
	{
		 /// <summary>
		 /// SupportByVersion MSComctlLib 6
		 /// </summary>
		 /// <remarks>0</remarks>
		 [SupportByVersion("MSComctlLib", 6)]
		 ccOLEDropEffectNone = 0,

		 /// <summary>
		 /// SupportByVersion MSComctlLib 6
		 /// </summary>
		 /// <remarks>1</remarks>
		 [SupportByVersion("MSComctlLib", 6)]
		 ccOLEDropEffectCopy = 1,

		 /// <summary>
		 /// SupportByVersion MSComctlLib 6
		 /// </summary>
		 /// <remarks>2</remarks>
		 [SupportByVersion("MSComctlLib", 6)]
		 ccOLEDropEffectMove = 2,

		 /// <summary>
		 /// SupportByVersion MSComctlLib 6
		 /// </summary>
		 /// <remarks>-2147483648</remarks>
		 [SupportByVersion("MSComctlLib", 6)]
		 ccOLEDropEffectScroll = -2147483648
	}
}