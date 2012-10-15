using System;
using NetOffice;
namespace NetOffice.MSProjectApi.Enums
{
	 /// <summary>
	 /// SupportByVersion MSProject 12, 14, 15
	 /// </summary>
	[SupportByVersionAttribute("MSProject", 12,14,15)]
	[EntityTypeAttribute(EntityType.IsEnum)]
	public enum PjRecalcDriverType
	{
		 /// <summary>
		 /// SupportByVersion MSProject 12, 14, 15
		 /// </summary>
		 /// <remarks>1</remarks>
		 [SupportByVersionAttribute("MSProject", 12,14,15)]
		 pjDriverActuals = 1,

		 /// <summary>
		 /// SupportByVersion MSProject 12, 14, 15
		 /// </summary>
		 /// <remarks>2</remarks>
		 [SupportByVersionAttribute("MSProject", 12,14,15)]
		 pjDriverLevelingDelay = 2,

		 /// <summary>
		 /// SupportByVersion MSProject 12, 14, 15
		 /// </summary>
		 /// <remarks>4</remarks>
		 [SupportByVersionAttribute("MSProject", 12,14,15)]
		 pjDriverConstraint = 4,

		 /// <summary>
		 /// SupportByVersion MSProject 12, 14, 15
		 /// </summary>
		 /// <remarks>8</remarks>
		 [SupportByVersionAttribute("MSProject", 12,14,15)]
		 pjDriverPredecessor = 8,

		 /// <summary>
		 /// SupportByVersion MSProject 12, 14, 15
		 /// </summary>
		 /// <remarks>16</remarks>
		 [SupportByVersionAttribute("MSProject", 12,14,15)]
		 pjDriverProjectStart = 16,

		 /// <summary>
		 /// SupportByVersion MSProject 12, 14, 15
		 /// </summary>
		 /// <remarks>32</remarks>
		 [SupportByVersionAttribute("MSProject", 12,14,15)]
		 pjDriverCalendar = 32,

		 /// <summary>
		 /// SupportByVersion MSProject 12, 14, 15
		 /// </summary>
		 /// <remarks>64</remarks>
		 [SupportByVersionAttribute("MSProject", 12,14,15)]
		 pjDriverChildTask = 64,

		 /// <summary>
		 /// SupportByVersion MSProject 12, 14, 15
		 /// </summary>
		 /// <remarks>128</remarks>
		 [SupportByVersionAttribute("MSProject", 12,14,15)]
		 pjDriverParentTask = 128,

		 /// <summary>
		 /// SupportByVersion MSProject 12, 14, 15
		 /// </summary>
		 /// <remarks>256</remarks>
		 [SupportByVersionAttribute("MSProject", 12,14,15)]
		 pjDriverDeadlineTask = 256
	}
}