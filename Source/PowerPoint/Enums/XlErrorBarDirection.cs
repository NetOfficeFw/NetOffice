using System;
using NetOffice;
using NetOffice.Attributes;
namespace NetOffice.PowerPointApi.Enums
{
	 /// <summary>
	 /// SupportByVersion PowerPoint 14, 15, 16
	 /// </summary>
	 ///<remarks> MSDN Online Documentation: <see href="https://docs.microsoft.com/en-us/office/vba/api/PowerPoint.XlErrorBarDirection"/> </remarks>
	[SupportByVersion("PowerPoint", 14,15,16)]
	[EntityType(EntityType.IsEnum)]
	public enum XlErrorBarDirection
	{
		 /// <summary>
		 /// SupportByVersion PowerPoint 14, 15, 16
		 /// </summary>
		 /// <remarks>-4168</remarks>
		 [SupportByVersion("PowerPoint", 14,15,16)]
		 xlChartX = -4168,

		 /// <summary>
		 /// SupportByVersion PowerPoint 14, 15, 16
		 /// </summary>
		 /// <remarks>1</remarks>
		 [SupportByVersion("PowerPoint", 14,15,16)]
		 xlChartY = 1
	}
}