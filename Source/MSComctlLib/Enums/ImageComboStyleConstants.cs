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
	public enum ImageComboStyleConstants
	{
		 /// <summary>
		 /// SupportByVersion MSComctlLib 6
		 /// </summary>
		 /// <remarks>0</remarks>
		 [SupportByVersion("MSComctlLib", 6)]
		 ImgCboDropdownCombo = 0,

		 /// <summary>
		 /// SupportByVersion MSComctlLib 6
		 /// </summary>
		 /// <remarks>1</remarks>
		 [SupportByVersion("MSComctlLib", 6)]
		 ImgCboSimpleCombo = 1,

		 /// <summary>
		 /// SupportByVersion MSComctlLib 6
		 /// </summary>
		 /// <remarks>2</remarks>
		 [SupportByVersion("MSComctlLib", 6)]
		 ImgCboDropdownList = 2
	}
}