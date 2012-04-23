using System;
using NetOffice;
namespace NetOffice.PowerPointApi.Enums
{
	 /// <summary>
	 /// SupportByVersion PowerPoint 14
	 /// </summary>
	[SupportByVersionAttribute("PowerPoint", 14)]
	[EntityTypeAttribute(EntityType.IsEnum)]
	public enum XlReadingOrder
	{
		 /// <summary>
		 /// SupportByVersion PowerPoint 14
		 /// </summary>
		 /// <remarks>-5002</remarks>
		 [SupportByVersionAttribute("PowerPoint", 14)]
		 xlContext = -5002,

		 /// <summary>
		 /// SupportByVersion PowerPoint 14
		 /// </summary>
		 /// <remarks>-5003</remarks>
		 [SupportByVersionAttribute("PowerPoint", 14)]
		 xlLTR = -5003,

		 /// <summary>
		 /// SupportByVersion PowerPoint 14
		 /// </summary>
		 /// <remarks>-5004</remarks>
		 [SupportByVersionAttribute("PowerPoint", 14)]
		 xlRTL = -5004
	}
}