using System;
using NetOffice;
namespace NetOffice.MSComctlLibApi.Enums
{
	 /// <summary>
	 /// SupportByVersion MSComctlLib 2
	 /// </summary>
	[SupportByVersionAttribute("MSComctlLib", 2)]
	[EntityTypeAttribute(EntityType.IsEnum)]
	public enum ImageDrawConstants
	{
		 /// <summary>
		 /// SupportByVersion MSComctlLib 2
		 /// </summary>
		 /// <remarks>0</remarks>
		 [SupportByVersionAttribute("MSComctlLib", 2)]
		 imlNormal = 0,

		 /// <summary>
		 /// SupportByVersion MSComctlLib 2
		 /// </summary>
		 /// <remarks>1</remarks>
		 [SupportByVersionAttribute("MSComctlLib", 2)]
		 imlTransparent = 1,

		 /// <summary>
		 /// SupportByVersion MSComctlLib 2
		 /// </summary>
		 /// <remarks>2</remarks>
		 [SupportByVersionAttribute("MSComctlLib", 2)]
		 imlSelected = 2,

		 /// <summary>
		 /// SupportByVersion MSComctlLib 2
		 /// </summary>
		 /// <remarks>3</remarks>
		 [SupportByVersionAttribute("MSComctlLib", 2)]
		 imlFocus = 3
	}
}