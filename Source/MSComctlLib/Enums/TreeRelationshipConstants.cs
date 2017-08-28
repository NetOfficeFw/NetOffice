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
	public enum TreeRelationshipConstants
	{
		 /// <summary>
		 /// SupportByVersion MSComctlLib 6
		 /// </summary>
		 /// <remarks>0</remarks>
		 [SupportByVersion("MSComctlLib", 6)]
		 tvwFirst = 0,

		 /// <summary>
		 /// SupportByVersion MSComctlLib 6
		 /// </summary>
		 /// <remarks>1</remarks>
		 [SupportByVersion("MSComctlLib", 6)]
		 tvwLast = 1,

		 /// <summary>
		 /// SupportByVersion MSComctlLib 6
		 /// </summary>
		 /// <remarks>2</remarks>
		 [SupportByVersion("MSComctlLib", 6)]
		 tvwNext = 2,

		 /// <summary>
		 /// SupportByVersion MSComctlLib 6
		 /// </summary>
		 /// <remarks>3</remarks>
		 [SupportByVersion("MSComctlLib", 6)]
		 tvwPrevious = 3,

		 /// <summary>
		 /// SupportByVersion MSComctlLib 6
		 /// </summary>
		 /// <remarks>4</remarks>
		 [SupportByVersion("MSComctlLib", 6)]
		 tvwChild = 4
	}
}