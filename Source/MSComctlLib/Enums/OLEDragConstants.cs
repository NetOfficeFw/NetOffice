using System;
using NetOffice;
namespace NetOffice.MSComctlLibApi.Enums
{
	 /// <summary>
	 /// SupportByVersion MSComctlLib 2
	 /// </summary>
	[SupportByVersionAttribute("MSComctlLib", 2)]
	[EntityTypeAttribute(EntityType.IsEnum)]
	public enum OLEDragConstants
	{
		 /// <summary>
		 /// SupportByVersion MSComctlLib 2
		 /// </summary>
		 /// <remarks>0</remarks>
		 [SupportByVersionAttribute("MSComctlLib", 2)]
		 ccOLEDragManual = 0,

		 /// <summary>
		 /// SupportByVersion MSComctlLib 2
		 /// </summary>
		 /// <remarks>1</remarks>
		 [SupportByVersionAttribute("MSComctlLib", 2)]
		 ccOLEDragAutomatic = 1
	}
}