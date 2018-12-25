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
	public enum PjAssignmentWarnings
	{
		 /// <summary>
		 /// SupportByVersion MSProject 16
		 /// </summary>
		 /// <remarks>1</remarks>
		 [SupportByVersion("MSProject", 16)]
		 pjAssignmentWarningNoOverlappingCommittedEngagement = 1,

		 /// <summary>
		 /// SupportByVersion MSProject 16
		 /// </summary>
		 /// <remarks>2</remarks>
		 [SupportByVersion("MSProject", 16)]
		 pjAssignmentWarningPartiallyOverlappingCommittedEngagement = 2,

		 /// <summary>
		 /// SupportByVersion MSProject 16
		 /// </summary>
		 /// <remarks>4</remarks>
		 [SupportByVersion("MSProject", 16)]
		 pjAssignmentWarningOverlappingWithLessWorkCommittedEngagement = 4,

		 /// <summary>
		 /// SupportByVersion MSProject 16
		 /// </summary>
		 /// <remarks>8</remarks>
		 [SupportByVersion("MSProject", 16)]
		 pjAssignmentWarningOverlappingProposedEngagement = 8,

		 /// <summary>
		 /// SupportByVersion MSProject 16
		 /// </summary>
		 /// <remarks>16</remarks>
		 [SupportByVersion("MSProject", 16)]
		 pjAssignmentWarningOverlappingDraftEngagement = 16
	}
}