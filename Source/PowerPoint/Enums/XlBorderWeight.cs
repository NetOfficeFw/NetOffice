using System;
using NetOffice;
using NetOffice.Attributes;
namespace NetOffice.PowerPointApi.Enums
{
	 /// <summary>
	 /// SupportByVersion PowerPoint 14, 15, 16
	 /// </summary>
	 ///<remarks> MSDN Online Documentation: <see href="https://docs.microsoft.com/en-us/office/vba/api/PowerPoint.XlBorderWeight"/> </remarks>
	[SupportByVersion("PowerPoint", 14,15,16)]
	[EntityType(EntityType.IsEnum)]
	public enum XlBorderWeight
	{
		 /// <summary>
		 /// SupportByVersion PowerPoint 14, 15, 16
		 /// </summary>
		 /// <remarks>1</remarks>
		 [SupportByVersion("PowerPoint", 14,15,16)]
		 xlHairline = 1,

		 /// <summary>
		 /// SupportByVersion PowerPoint 14, 15, 16
		 /// </summary>
		 /// <remarks>-4138</remarks>
		 [SupportByVersion("PowerPoint", 14,15,16)]
		 xlMedium = -4138,

		 /// <summary>
		 /// SupportByVersion PowerPoint 14, 15, 16
		 /// </summary>
		 /// <remarks>4</remarks>
		 [SupportByVersion("PowerPoint", 14,15,16)]
		 xlThick = 4,

		 /// <summary>
		 /// SupportByVersion PowerPoint 14, 15, 16
		 /// </summary>
		 /// <remarks>2</remarks>
		 [SupportByVersion("PowerPoint", 14,15,16)]
		 xlThin = 2
	}
}