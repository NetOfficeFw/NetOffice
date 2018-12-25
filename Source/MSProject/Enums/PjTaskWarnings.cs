using System;
using NetOffice;
using NetOffice.Attributes;
namespace NetOffice.MSProjectApi.Enums
{
	 /// <summary>
	 /// SupportByVersion MSProject 11, 14, 16
	 /// </summary>
	 ///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff860436(v=office.14).aspx </remarks>
	[SupportByVersion("MSProject", 11,14,16)]
	[EntityType(EntityType.IsEnum)]
	public enum PjTaskWarnings
	{
		 /// <summary>
		 /// SupportByVersion MSProject 11, 14, 16
		 /// </summary>
		 /// <remarks>1</remarks>
		 [SupportByVersion("MSProject", 11,14,16)]
		 pjTaskWarningShadowFinishesLaterDueToLink = 1,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 14, 16
		 /// </summary>
		 /// <remarks>2</remarks>
		 [SupportByVersion("MSProject", 11,14,16)]
		 pjTaskWarningShadowFinishesEarlierDueToLink = 2,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 14, 16
		 /// </summary>
		 /// <remarks>4</remarks>
		 [SupportByVersion("MSProject", 11,14,16)]
		 pjTaskWarningSubTaskStartingBeforeParentStart = 4,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 14, 16
		 /// </summary>
		 /// <remarks>8</remarks>
		 [SupportByVersion("MSProject", 11,14,16)]
		 pjTaskWarningSubTaskStartingAfterParentStart = 8,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 14, 16
		 /// </summary>
		 /// <remarks>16</remarks>
		 [SupportByVersion("MSProject", 11,14,16)]
		 pjTaskWarningSubTaskFinishingAfterParentFinish = 16,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 14, 16
		 /// </summary>
		 /// <remarks>32</remarks>
		 [SupportByVersion("MSProject", 11,14,16)]
		 pjTaskWarningSummaryInconsistentStart = 32,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 14, 16
		 /// </summary>
		 /// <remarks>64</remarks>
		 [SupportByVersion("MSProject", 11,14,16)]
		 pjTaskWarningResourceBeyondMaxUnit = 64,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 14, 16
		 /// </summary>
		 /// <remarks>128</remarks>
		 [SupportByVersion("MSProject", 11,14,16)]
		 pjTaskWarningResourceOverallocated = 128,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 14, 16
		 /// </summary>
		 /// <remarks>256</remarks>
		 [SupportByVersion("MSProject", 11,14,16)]
		 pjTaskWarningShadowIncorrectByConstraintOnly = 256,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 14, 16
		 /// </summary>
		 /// <remarks>512</remarks>
		 [SupportByVersion("MSProject", 11,14,16)]
		 pjTaskWarningShadowIncorrectByLevelingDelayOnly = 512,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 14, 16
		 /// </summary>
		 /// <remarks>1024</remarks>
		 [SupportByVersion("MSProject", 11,14,16)]
		 pjTaskWarningShadowDateDifferent = 1024,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 14, 16
		 /// </summary>
		 /// <remarks>2048</remarks>
		 [SupportByVersion("MSProject", 11,14,16)]
		 pjTaskWarningSummaryInconsistentFinish = 2048,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 14, 16
		 /// </summary>
		 /// <remarks>4096</remarks>
		 [SupportByVersion("MSProject", 11,14,16)]
		 pjTaskWarningTaskStartingInNonWorkingTime = 4096,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 14, 16
		 /// </summary>
		 /// <remarks>8192</remarks>
		 [SupportByVersion("MSProject", 11,14,16)]
		 pjTaskWarningTaskFinishingInNonWorkingTime = 8192,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 14, 16
		 /// </summary>
		 /// <remarks>16384</remarks>
		 [SupportByVersion("MSProject", 11,14,16)]
		 pjTaskWarningAssnOverallocatedInNonWorkingTime = 16384,

		 /// <summary>
		 /// SupportByVersion MSProject 16
		 /// </summary>
		 /// <remarks>32768</remarks>
		 [SupportByVersion("MSProject", 16)]
		 pjTaskWarningAssignmentEngagementViolation = 32768,

		 /// <summary>
		 /// SupportByVersion MSProject 16
		 /// </summary>
		 /// <remarks>65536</remarks>
		 [SupportByVersion("MSProject", 16)]
		 pjTaskWarningAssignmentWorkingInProposedEngagedTime = 65536,

		 /// <summary>
		 /// SupportByVersion MSProject 16
		 /// </summary>
		 /// <remarks>131072</remarks>
		 [SupportByVersion("MSProject", 16)]
		 pjTaskWarningAssignmentWorkingInDraftEngagedTime = 131072,

		 /// <summary>
		 /// SupportByVersion MSProject 16
		 /// </summary>
		 /// <remarks>262144</remarks>
		 [SupportByVersion("MSProject", 16)]
		 pjTaskWarningResourceEngagementViolation = 262144,

		 /// <summary>
		 /// SupportByVersion MSProject 16
		 /// </summary>
		 /// <remarks>524288</remarks>
		 [SupportByVersion("MSProject", 16)]
		 pjTaskWarningResourceWorkingInProposedEngagedTime = 524288,

		 /// <summary>
		 /// SupportByVersion MSProject 16
		 /// </summary>
		 /// <remarks>1048576</remarks>
		 [SupportByVersion("MSProject", 16)]
		 pjTaskWarningResourceWorkingInDraftEngagedTime = 1048576
	}
}