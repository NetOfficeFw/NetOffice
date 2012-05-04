using System;
using NetOffice;
namespace NetOffice.MSComctlLibApi.Enums
{
	 /// <summary>
	 /// SupportByVersion MSComctlLib 2
	 /// </summary>
	[SupportByVersionAttribute("MSComctlLib", 2)]
	[EntityTypeAttribute(EntityType.IsEnum)]
	public enum OLEDropEffectConstants
	{
		 /// <summary>
		 /// SupportByVersion MSComctlLib 2
		 /// </summary>
		 /// <remarks>0</remarks>
		 [SupportByVersionAttribute("MSComctlLib", 2)]
		 ccOLEDropEffectNone = 0,

		 /// <summary>
		 /// SupportByVersion MSComctlLib 2
		 /// </summary>
		 /// <remarks>1</remarks>
		 [SupportByVersionAttribute("MSComctlLib", 2)]
		 ccOLEDropEffectCopy = 1,

		 /// <summary>
		 /// SupportByVersion MSComctlLib 2
		 /// </summary>
		 /// <remarks>2</remarks>
		 [SupportByVersionAttribute("MSComctlLib", 2)]
		 ccOLEDropEffectMove = 2,

		 /// <summary>
		 /// SupportByVersion MSComctlLib 2
		 /// </summary>
		 /// <remarks>-2147483648</remarks>
		 [SupportByVersionAttribute("MSComctlLib", 2)]
		 ccOLEDropEffectScroll = -2147483648
	}
}