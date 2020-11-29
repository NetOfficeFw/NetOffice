using System;
using NetOffice;
using NetOffice.Attributes;
namespace NetOffice.PowerPointApi.Enums
{
	 /// <summary>
	 /// SupportByVersion PowerPoint 12, 14, 15, 16
	 /// </summary>
	 ///<remarks> MSDN Online Documentation: <see href="https://docs.microsoft.com/en-us/office/vba/api/PowerPoint.PpCheckInVersionType"/> </remarks>
	[SupportByVersion("PowerPoint", 12,14,15,16)]
	[EntityType(EntityType.IsEnum)]
	public enum PpCheckInVersionType
	{
		 /// <summary>
		 /// SupportByVersion PowerPoint 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>0</remarks>
		 [SupportByVersion("PowerPoint", 12,14,15,16)]
		 ppCheckInMinorVersion = 0,

		 /// <summary>
		 /// SupportByVersion PowerPoint 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>1</remarks>
		 [SupportByVersion("PowerPoint", 12,14,15,16)]
		 ppCheckInMajorVersion = 1,

		 /// <summary>
		 /// SupportByVersion PowerPoint 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>2</remarks>
		 [SupportByVersion("PowerPoint", 12,14,15,16)]
		 ppCheckInOverwriteVersion = 2
	}
}