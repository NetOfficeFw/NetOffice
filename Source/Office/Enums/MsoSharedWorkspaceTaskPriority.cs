using System;
using NetOffice;
using NetOffice.Attributes;
namespace NetOffice.OfficeApi.Enums
{
	 /// <summary>
	 /// SupportByVersion Office 11, 12, 14, 15, 16
	 /// </summary>
	 ///<remarks> MSDN Online Documentation: <see href="https://docs.microsoft.com/en-us/office/vba/api/Office.MsoSharedWorkspaceTaskPriority"/> </remarks>
	[SupportByVersion("Office", 11,12,14,15,16)]
	[EntityType(EntityType.IsEnum)]
	public enum MsoSharedWorkspaceTaskPriority
	{
		 /// <summary>
		 /// SupportByVersion Office 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>1</remarks>
		 [SupportByVersion("Office", 11,12,14,15,16)]
		 msoSharedWorkspaceTaskPriorityHigh = 1,

		 /// <summary>
		 /// SupportByVersion Office 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>2</remarks>
		 [SupportByVersion("Office", 11,12,14,15,16)]
		 msoSharedWorkspaceTaskPriorityNormal = 2,

		 /// <summary>
		 /// SupportByVersion Office 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>3</remarks>
		 [SupportByVersion("Office", 11,12,14,15,16)]
		 msoSharedWorkspaceTaskPriorityLow = 3
	}
}