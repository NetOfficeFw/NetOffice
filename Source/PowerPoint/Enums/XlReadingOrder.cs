using System;
using NetOffice;
namespace NetOffice.PowerPointApi.Enums
{
	 /// <summary>
	 /// SupportByVersion PowerPoint 14, 15
	 /// </summary>
	[SupportByVersionAttribute("PowerPoint", 14,15)]
	[EntityTypeAttribute(EntityType.IsEnum)]
	public enum XlReadingOrder
	{
		 /// <summary>
		 /// SupportByVersion PowerPoint 14, 15
		 /// </summary>
		 /// <remarks>-5002</remarks>
		 [SupportByVersionAttribute("PowerPoint", 14,15)]
		 xlContext = -5002,

		 /// <summary>
		 /// SupportByVersion PowerPoint 14, 15
		 /// </summary>
		 /// <remarks>-5003</remarks>
		 [SupportByVersionAttribute("PowerPoint", 14,15)]
		 xlLTR = -5003,

		 /// <summary>
		 /// SupportByVersion PowerPoint 14, 15
		 /// </summary>
		 /// <remarks>-5004</remarks>
		 [SupportByVersionAttribute("PowerPoint", 14,15)]
		 xlRTL = -5004
	}
}