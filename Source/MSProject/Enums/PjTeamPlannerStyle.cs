using System;
using NetOffice;
using NetOffice.Attributes;
namespace NetOffice.MSProjectApi.Enums
{
	 /// <summary>
	 /// SupportByVersion MSProject 11, 14, 16
	 /// </summary>
	 ///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff864349(v=office.14).aspx </remarks>
	[SupportByVersion("MSProject", 11,14,16)]
	[EntityType(EntityType.IsEnum)]
	public enum PjTeamPlannerStyle
	{
		 /// <summary>
		 /// SupportByVersion MSProject 11, 14, 16
		 /// </summary>
		 /// <remarks>0</remarks>
		 [SupportByVersion("MSProject", 11,14,16)]
		 pjTPScheduledWork = 0,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 14, 16
		 /// </summary>
		 /// <remarks>1</remarks>
		 [SupportByVersion("MSProject", 11,14,16)]
		 pjTPActualWork = 1,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 14, 16
		 /// </summary>
		 /// <remarks>2</remarks>
		 [SupportByVersion("MSProject", 11,14,16)]
		 pjTPSRA = 2,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 14, 16
		 /// </summary>
		 /// <remarks>3</remarks>
		 [SupportByVersion("MSProject", 11,14,16)]
		 pjTPLateTask = 3,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 14, 16
		 /// </summary>
		 /// <remarks>4</remarks>
		 [SupportByVersion("MSProject", 11,14,16)]
		 pjTPManualTask = 4
	}
}