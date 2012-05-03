using System;
using NetOffice;
namespace NetOffice.PowerPointApi.Enums
{
	 /// <summary>
	 /// SupportByVersion PowerPoint 12, 14
	 /// </summary>
	[SupportByVersionAttribute("PowerPoint", 12,14)]
	[EntityTypeAttribute(EntityType.IsEnum)]
	public enum PpFixedFormatType
	{
		 /// <summary>
		 /// SupportByVersion PowerPoint 12, 14
		 /// </summary>
		 /// <remarks>1</remarks>
		 [SupportByVersionAttribute("PowerPoint", 12,14)]
		 ppFixedFormatTypeXPS = 1,

		 /// <summary>
		 /// SupportByVersion PowerPoint 12, 14
		 /// </summary>
		 /// <remarks>2</remarks>
		 [SupportByVersionAttribute("PowerPoint", 12,14)]
		 ppFixedFormatTypePDF = 2
	}
}