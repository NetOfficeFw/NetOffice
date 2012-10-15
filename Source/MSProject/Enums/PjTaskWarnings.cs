using System;
using NetOffice;
namespace NetOffice.MSProjectApi.Enums
{
	 /// <summary>
	 /// SupportByVersion MSProject 14, 15
	 /// </summary>
	[SupportByVersionAttribute("MSProject", 14,15)]
	[EntityTypeAttribute(EntityType.IsEnum)]
	public enum PjTaskWarnings
	{
		 /// <summary>
		 /// SupportByVersion MSProject 14, 15
		 /// </summary>
		 /// <remarks>1</remarks>
		 [SupportByVersionAttribute("MSProject", 14,15)]
		 pjTaskWarningShadowFinishesLaterDueToLink = 1,

		 /// <summary>
		 /// SupportByVersion MSProject 14, 15
		 /// </summary>
		 /// <remarks>2</remarks>
		 [SupportByVersionAttribute("MSProject", 14,15)]
		 pjTaskWarningShadowFinishesEarlierDueToLink = 2,

		 /// <summary>
		 /// SupportByVersion MSProject 14, 15
		 /// </summary>
		 /// <remarks>4</remarks>
		 [SupportByVersionAttribute("MSProject", 14,15)]
		 pjTaskWarningSubTaskStartingBeforeParentStart = 4,

		 /// <summary>
		 /// SupportByVersion MSProject 14, 15
		 /// </summary>
		 /// <remarks>8</remarks>
		 [SupportByVersionAttribute("MSProject", 14,15)]
		 pjTaskWarningSubTaskStartingAfterParentStart = 8,

		 /// <summary>
		 /// SupportByVersion MSProject 14, 15
		 /// </summary>
		 /// <remarks>16</remarks>
		 [SupportByVersionAttribute("MSProject", 14,15)]
		 pjTaskWarningSubTaskFinishingAfterParentFinish = 16,

		 /// <summary>
		 /// SupportByVersion MSProject 14, 15
		 /// </summary>
		 /// <remarks>32</remarks>
		 [SupportByVersionAttribute("MSProject", 14,15)]
		 pjTaskWarningSummaryInconsistentStart = 32,

		 /// <summary>
		 /// SupportByVersion MSProject 14, 15
		 /// </summary>
		 /// <remarks>64</remarks>
		 [SupportByVersionAttribute("MSProject", 14,15)]
		 pjTaskWarningResourceBeyondMaxUnit = 64,

		 /// <summary>
		 /// SupportByVersion MSProject 14, 15
		 /// </summary>
		 /// <remarks>128</remarks>
		 [SupportByVersionAttribute("MSProject", 14,15)]
		 pjTaskWarningResourceOverallocated = 128,

		 /// <summary>
		 /// SupportByVersion MSProject 14, 15
		 /// </summary>
		 /// <remarks>256</remarks>
		 [SupportByVersionAttribute("MSProject", 14,15)]
		 pjTaskWarningShadowIncorrectByConstraintOnly = 256,

		 /// <summary>
		 /// SupportByVersion MSProject 14, 15
		 /// </summary>
		 /// <remarks>512</remarks>
		 [SupportByVersionAttribute("MSProject", 14,15)]
		 pjTaskWarningShadowIncorrectByLevelingDelayOnly = 512,

		 /// <summary>
		 /// SupportByVersion MSProject 14, 15
		 /// </summary>
		 /// <remarks>1024</remarks>
		 [SupportByVersionAttribute("MSProject", 14,15)]
		 pjTaskWarningShadowDateDifferent = 1024,

		 /// <summary>
		 /// SupportByVersion MSProject 14, 15
		 /// </summary>
		 /// <remarks>2048</remarks>
		 [SupportByVersionAttribute("MSProject", 14,15)]
		 pjTaskWarningSummaryInconsistentFinish = 2048,

		 /// <summary>
		 /// SupportByVersion MSProject 14, 15
		 /// </summary>
		 /// <remarks>4096</remarks>
		 [SupportByVersionAttribute("MSProject", 14,15)]
		 pjTaskWarningTaskStartingInNonWorkingTime = 4096,

		 /// <summary>
		 /// SupportByVersion MSProject 14, 15
		 /// </summary>
		 /// <remarks>8192</remarks>
		 [SupportByVersionAttribute("MSProject", 14,15)]
		 pjTaskWarningTaskFinishingInNonWorkingTime = 8192,

		 /// <summary>
		 /// SupportByVersion MSProject 14, 15
		 /// </summary>
		 /// <remarks>16384</remarks>
		 [SupportByVersionAttribute("MSProject", 14,15)]
		 pjTaskWarningAssnOverallocatedInNonWorkingTime = 16384
	}
}