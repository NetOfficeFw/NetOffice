using System;
using NetOffice;
using NetOffice.Attributes;
namespace NetOffice.PowerPointApi.Enums
{
	 /// <summary>
	 /// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
	 /// </summary>
	 ///<remarks> MSDN Online Documentation: <see href="https://docs.microsoft.com/en-us/office/vba/api/PowerPoint.PpArrangeStyle"/> </remarks>
	[SupportByVersion("PowerPoint", 9,10,11,12,14,15,16)]
	[EntityType(EntityType.IsEnum)]
	public enum PpArrangeStyle
	{
		 /// <summary>
		 /// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>1</remarks>
		 [SupportByVersion("PowerPoint", 9,10,11,12,14,15,16)]
		 ppArrangeTiled = 1,

		 /// <summary>
		 /// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>2</remarks>
		 [SupportByVersion("PowerPoint", 9,10,11,12,14,15,16)]
		 ppArrangeCascade = 2
	}
}