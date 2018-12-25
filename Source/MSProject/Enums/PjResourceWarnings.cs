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
	public enum PjResourceWarnings
	{
		 /// <summary>
		 /// SupportByVersion MSProject 16
		 /// </summary>
		 /// <remarks>1</remarks>
		 [SupportByVersion("MSProject", 16)]
		 pjResourceWarningAssignmentEngagementViolation = 1,

		 /// <summary>
		 /// SupportByVersion MSProject 16
		 /// </summary>
		 /// <remarks>2</remarks>
		 [SupportByVersion("MSProject", 16)]
		 pjResourceWarningAssignmentWorkingInProposedEngagedTime = 2,

		 /// <summary>
		 /// SupportByVersion MSProject 16
		 /// </summary>
		 /// <remarks>4</remarks>
		 [SupportByVersion("MSProject", 16)]
		 pjResourceWarningAssignmentWorkingInDraftEngagedTime = 4,

		 /// <summary>
		 /// SupportByVersion MSProject 16
		 /// </summary>
		 /// <remarks>8</remarks>
		 [SupportByVersion("MSProject", 16)]
		 pjResourceWarningEngagementViolation = 8,

		 /// <summary>
		 /// SupportByVersion MSProject 16
		 /// </summary>
		 /// <remarks>16</remarks>
		 [SupportByVersion("MSProject", 16)]
		 pjResourceWarningWorkingInProposedEngagedTime = 16,

		 /// <summary>
		 /// SupportByVersion MSProject 16
		 /// </summary>
		 /// <remarks>32</remarks>
		 [SupportByVersion("MSProject", 16)]
		 pjResourceWarningWorkingInDraftEngagedTime = 32
	}
}