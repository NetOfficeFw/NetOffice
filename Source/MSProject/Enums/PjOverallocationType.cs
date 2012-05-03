using System;
using NetOffice;
namespace NetOffice.MSProjectApi.Enums
{
	 /// <summary>
	 /// SupportByVersion MSProject 14
	 /// </summary>
	[SupportByVersionAttribute("MSProject", 14)]
	[EntityTypeAttribute(EntityType.IsEnum)]
	public enum PjOverallocationType
	{
		 /// <summary>
		 /// SupportByVersion MSProject 14
		 /// </summary>
		 /// <remarks>0</remarks>
		 [SupportByVersionAttribute("MSProject", 14)]
		 pjOverallocationTypeNone = 0,

		 /// <summary>
		 /// SupportByVersion MSProject 14
		 /// </summary>
		 /// <remarks>1</remarks>
		 [SupportByVersionAttribute("MSProject", 14)]
		 pjOverallocationTypeAboveMaxUnits = 1,

		 /// <summary>
		 /// SupportByVersion MSProject 14
		 /// </summary>
		 /// <remarks>2</remarks>
		 [SupportByVersionAttribute("MSProject", 14)]
		 pjOverallocationTypeWorkingInNonWorkingTime = 2,

		 /// <summary>
		 /// SupportByVersion MSProject 14
		 /// </summary>
		 /// <remarks>3</remarks>
		 [SupportByVersionAttribute("MSProject", 14)]
		 pjOverallocationTypeWorkingOnOtherTasks = 3
	}
}