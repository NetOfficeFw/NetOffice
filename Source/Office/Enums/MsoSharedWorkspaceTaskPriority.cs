using System;
using NetOffice;
namespace NetOffice.OfficeApi.Enums
{
	 /// <summary>
	 /// SupportByVersion Office 11, 12, 14, 15
	 /// </summary>
	[SupportByVersionAttribute("Office", 11,12,14,15)]
	[EntityTypeAttribute(EntityType.IsEnum)]
	public enum MsoSharedWorkspaceTaskPriority
	{
		 /// <summary>
		 /// SupportByVersion Office 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>1</remarks>
		 [SupportByVersionAttribute("Office", 11,12,14,15)]
		 msoSharedWorkspaceTaskPriorityHigh = 1,

		 /// <summary>
		 /// SupportByVersion Office 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>2</remarks>
		 [SupportByVersionAttribute("Office", 11,12,14,15)]
		 msoSharedWorkspaceTaskPriorityNormal = 2,

		 /// <summary>
		 /// SupportByVersion Office 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>3</remarks>
		 [SupportByVersionAttribute("Office", 11,12,14,15)]
		 msoSharedWorkspaceTaskPriorityLow = 3
	}
}