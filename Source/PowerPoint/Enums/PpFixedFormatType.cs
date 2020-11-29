using System;
using NetOffice;
using NetOffice.Attributes;
namespace NetOffice.PowerPointApi.Enums
{
	 /// <summary>
	 /// SupportByVersion PowerPoint 12, 14, 15, 16
	 /// </summary>
	 ///<remarks> MSDN Online Documentation: <see href="https://docs.microsoft.com/en-us/office/vba/api/PowerPoint.PpFixedFormatType"/> </remarks>
	[SupportByVersion("PowerPoint", 12,14,15,16)]
	[EntityType(EntityType.IsEnum)]
	public enum PpFixedFormatType
	{
		 /// <summary>
		 /// SupportByVersion PowerPoint 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>1</remarks>
		 [SupportByVersion("PowerPoint", 12,14,15,16)]
		 ppFixedFormatTypeXPS = 1,

		 /// <summary>
		 /// SupportByVersion PowerPoint 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>2</remarks>
		 [SupportByVersion("PowerPoint", 12,14,15,16)]
		 ppFixedFormatTypePDF = 2
	}
}