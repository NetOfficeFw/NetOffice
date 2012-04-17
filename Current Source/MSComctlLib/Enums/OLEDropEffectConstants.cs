using System;
using LateBindingApi.Core;
namespace NetOffice.MSComctlLibApi.Enums
{
	 /// <summary>
	 /// SupportByVersion MSComctlLib 6.0
	 /// </summary>
	[SupportByVersionAttribute("MSComctlLib", 6.0)]
	[EntityTypeAttribute(EntityType.IsEnum)]
	public enum OLEDropEffectConstants
	{
		 /// <summary>
		 /// SupportByVersion MSComctlLib 6.0
		 /// </summary>
		 /// <remarks>0</remarks>
		 [SupportByVersionAttribute("MSComctlLib", 6.0)]
		 ccOLEDropEffectNone = 0,

		 /// <summary>
		 /// SupportByVersion MSComctlLib 6.0
		 /// </summary>
		 /// <remarks>1</remarks>
		 [SupportByVersionAttribute("MSComctlLib", 6.0)]
		 ccOLEDropEffectCopy = 1,

		 /// <summary>
		 /// SupportByVersion MSComctlLib 6.0
		 /// </summary>
		 /// <remarks>2</remarks>
		 [SupportByVersionAttribute("MSComctlLib", 6.0)]
		 ccOLEDropEffectMove = 2,

		 /// <summary>
		 /// SupportByVersion MSComctlLib 6.0
		 /// </summary>
		 /// <remarks>-2147483648</remarks>
		 [SupportByVersionAttribute("MSComctlLib", 6.0)]
		 ccOLEDropEffectScroll = -2147483648
	}
}