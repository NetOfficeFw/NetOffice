using System;
using NetOffice;
namespace NetOffice.PowerPointApi.Enums
{
	 /// <summary>
	 /// SupportByVersion PowerPoint 12, 14, 15
	 /// </summary>
	 ///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff746243.aspx </remarks>
	[SupportByVersionAttribute("PowerPoint", 12,14,15)]
	[EntityTypeAttribute(EntityType.IsEnum)]
	public enum PpCheckInVersionType
	{
		 /// <summary>
		 /// SupportByVersion PowerPoint 12, 14, 15
		 /// </summary>
		 /// <remarks>0</remarks>
		 [SupportByVersionAttribute("PowerPoint", 12,14,15)]
		 ppCheckInMinorVersion = 0,

		 /// <summary>
		 /// SupportByVersion PowerPoint 12, 14, 15
		 /// </summary>
		 /// <remarks>1</remarks>
		 [SupportByVersionAttribute("PowerPoint", 12,14,15)]
		 ppCheckInMajorVersion = 1,

		 /// <summary>
		 /// SupportByVersion PowerPoint 12, 14, 15
		 /// </summary>
		 /// <remarks>2</remarks>
		 [SupportByVersionAttribute("PowerPoint", 12,14,15)]
		 ppCheckInOverwriteVersion = 2
	}
}