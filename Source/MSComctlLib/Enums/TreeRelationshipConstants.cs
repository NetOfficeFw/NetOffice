using System;
using NetOffice;
namespace NetOffice.MSComctlLibApi.Enums
{
	 /// <summary>
	 /// SupportByVersion MSComctlLib 2
	 /// </summary>
	[SupportByVersionAttribute("MSComctlLib", 2)]
	[EntityTypeAttribute(EntityType.IsEnum)]
	public enum TreeRelationshipConstants
	{
		 /// <summary>
		 /// SupportByVersion MSComctlLib 2
		 /// </summary>
		 /// <remarks>0</remarks>
		 [SupportByVersionAttribute("MSComctlLib", 2)]
		 tvwFirst = 0,

		 /// <summary>
		 /// SupportByVersion MSComctlLib 2
		 /// </summary>
		 /// <remarks>1</remarks>
		 [SupportByVersionAttribute("MSComctlLib", 2)]
		 tvwLast = 1,

		 /// <summary>
		 /// SupportByVersion MSComctlLib 2
		 /// </summary>
		 /// <remarks>2</remarks>
		 [SupportByVersionAttribute("MSComctlLib", 2)]
		 tvwNext = 2,

		 /// <summary>
		 /// SupportByVersion MSComctlLib 2
		 /// </summary>
		 /// <remarks>3</remarks>
		 [SupportByVersionAttribute("MSComctlLib", 2)]
		 tvwPrevious = 3,

		 /// <summary>
		 /// SupportByVersion MSComctlLib 2
		 /// </summary>
		 /// <remarks>4</remarks>
		 [SupportByVersionAttribute("MSComctlLib", 2)]
		 tvwChild = 4
	}
}