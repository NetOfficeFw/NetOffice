using System;
using LateBindingApi.Core;
namespace NetOffice.MSComctlLibApi.Enums
{
	 /// <summary>
	 /// SupportByVersion MSComctlLib 6.0
	 /// </summary>
	[SupportByVersionAttribute("MSComctlLib", 6.0)]
	[EntityTypeAttribute(EntityType.IsEnum)]
	public enum OrientationConstants
	{
		 /// <summary>
		 /// SupportByVersion MSComctlLib 6.0
		 /// </summary>
		 /// <remarks>0</remarks>
		 [SupportByVersionAttribute("MSComctlLib", 6.0)]
		 ccOrientationHorizontal = 0,

		 /// <summary>
		 /// SupportByVersion MSComctlLib 6.0
		 /// </summary>
		 /// <remarks>1</remarks>
		 [SupportByVersionAttribute("MSComctlLib", 6.0)]
		 ccOrientationVertical = 1
	}
}