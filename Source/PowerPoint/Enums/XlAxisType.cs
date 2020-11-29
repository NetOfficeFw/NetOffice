using System;
using NetOffice;
using NetOffice.Attributes;
namespace NetOffice.PowerPointApi.Enums
{
	 /// <summary>
	 /// SupportByVersion PowerPoint 14, 15, 16
	 /// </summary>
	 ///<remarks> MSDN Online Documentation: <see href="https://docs.microsoft.com/en-us/office/vba/api/PowerPoint.XlAxisType"/> </remarks>
	[SupportByVersion("PowerPoint", 14,15,16)]
	[EntityType(EntityType.IsEnum)]
	public enum XlAxisType
	{
		 /// <summary>
		 /// SupportByVersion PowerPoint 14, 15, 16
		 /// </summary>
		 /// <remarks>1</remarks>
		 [SupportByVersion("PowerPoint", 14,15,16)]
		 xlCategory = 1,

		 /// <summary>
		 /// SupportByVersion PowerPoint 14, 15, 16
		 /// </summary>
		 /// <remarks>3</remarks>
		 [SupportByVersion("PowerPoint", 14,15,16)]
		 xlSeriesAxis = 3,

		 /// <summary>
		 /// SupportByVersion PowerPoint 14, 15, 16
		 /// </summary>
		 /// <remarks>2</remarks>
		 [SupportByVersion("PowerPoint", 14,15,16)]
		 xlValue = 2
	}
}