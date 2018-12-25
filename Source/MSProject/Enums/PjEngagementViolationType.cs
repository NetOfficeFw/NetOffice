using System;
using NetOffice;
using NetOffice.Attributes;
namespace NetOffice.MSProjectApi.Enums
{
	 /// <summary>
	 /// SupportByVersion MSProject 16
	 /// </summary>
	[SupportByVersion("MSProject", 16)]
	[EntityType(EntityType.IsEnum)]
	public enum PjEngagementViolationType
	{
		 /// <summary>
		 /// SupportByVersion MSProject 16
		 /// </summary>
		 /// <remarks>0</remarks>
		 [SupportByVersion("MSProject", 16)]
		 pjEngagementViolationTypeNone = 0,

		 /// <summary>
		 /// SupportByVersion MSProject 16
		 /// </summary>
		 /// <remarks>1</remarks>
		 [SupportByVersion("MSProject", 16)]
		 pjEngagementViolationTypeTaskAssignmentEngagementViolation = 1,

		 /// <summary>
		 /// SupportByVersion MSProject 16
		 /// </summary>
		 /// <remarks>2</remarks>
		 [SupportByVersion("MSProject", 16)]
		 pjEngagementViolationTypeTaskAssignmentWorkingInProposedEngagedTime = 2,

		 /// <summary>
		 /// SupportByVersion MSProject 16
		 /// </summary>
		 /// <remarks>4</remarks>
		 [SupportByVersion("MSProject", 16)]
		 pjEngagementViolationTypeTaskAssignmentWorkingInDraftEngagedTime = 4,

		 /// <summary>
		 /// SupportByVersion MSProject 16
		 /// </summary>
		 /// <remarks>8</remarks>
		 [SupportByVersion("MSProject", 16)]
		 pjEngagementViolationTypeTaskResourceEngagementViolation = 8,

		 /// <summary>
		 /// SupportByVersion MSProject 16
		 /// </summary>
		 /// <remarks>16</remarks>
		 [SupportByVersion("MSProject", 16)]
		 pjEngagementViolationTypeTaskResourceWorkingInProposedEngagedTime = 16,

		 /// <summary>
		 /// SupportByVersion MSProject 16
		 /// </summary>
		 /// <remarks>32</remarks>
		 [SupportByVersion("MSProject", 16)]
		 pjEngagementViolationTypeTaskResourceWorkingInDraftEngagedTime = 32,

		 /// <summary>
		 /// SupportByVersion MSProject 16
		 /// </summary>
		 /// <remarks>64</remarks>
		 [SupportByVersion("MSProject", 16)]
		 pjEngagementViolationTypeResourceAssignmentEngagementViolation = 64,

		 /// <summary>
		 /// SupportByVersion MSProject 16
		 /// </summary>
		 /// <remarks>128</remarks>
		 [SupportByVersion("MSProject", 16)]
		 pjEngagementViolationTypeResourceAssignmentWorkingInProposedEngagedTime = 128,

		 /// <summary>
		 /// SupportByVersion MSProject 16
		 /// </summary>
		 /// <remarks>256</remarks>
		 [SupportByVersion("MSProject", 16)]
		 pjEngagementViolationTypeResourceAssignmentWorkingInDraftEngagedTime = 256,

		 /// <summary>
		 /// SupportByVersion MSProject 16
		 /// </summary>
		 /// <remarks>512</remarks>
		 [SupportByVersion("MSProject", 16)]
		 pjEngagementViolationTypeResourceEngagementViolation = 512,

		 /// <summary>
		 /// SupportByVersion MSProject 16
		 /// </summary>
		 /// <remarks>1024</remarks>
		 [SupportByVersion("MSProject", 16)]
		 pjEngagementViolationTypeResourceWorkingInProposedEngagedTime = 1024,

		 /// <summary>
		 /// SupportByVersion MSProject 16
		 /// </summary>
		 /// <remarks>2048</remarks>
		 [SupportByVersion("MSProject", 16)]
		 pjEngagementViolationTypeResourceWorkingInDraftEngagedTime = 2048,

		 /// <summary>
		 /// SupportByVersion MSProject 16
		 /// </summary>
		 /// <remarks>4096</remarks>
		 [SupportByVersion("MSProject", 16)]
		 pjEngagementViolationTypeAssignmentWorkingOutsideEngagedTime = 4096,

		 /// <summary>
		 /// SupportByVersion MSProject 16
		 /// </summary>
		 /// <remarks>8192</remarks>
		 [SupportByVersion("MSProject", 16)]
		 pjEngagementViolationTypeAssignmentWorkingPartiallyOutsideEngagedTime = 8192,

		 /// <summary>
		 /// SupportByVersion MSProject 16
		 /// </summary>
		 /// <remarks>16384</remarks>
		 [SupportByVersion("MSProject", 16)]
		 pjEngagementViolationTypeAssignmentWorkingAboveEngagedTime = 16384,

		 /// <summary>
		 /// SupportByVersion MSProject 16
		 /// </summary>
		 /// <remarks>32768</remarks>
		 [SupportByVersion("MSProject", 16)]
		 pjEngagementViolationTypeAssignmentWorkingInProposedEngagedTime = 32768,

		 /// <summary>
		 /// SupportByVersion MSProject 16
		 /// </summary>
		 /// <remarks>65536</remarks>
		 [SupportByVersion("MSProject", 16)]
		 pjEngagementViolationTypeAssignmentWorkingInDraftEngagedTime = 65536
	}
}