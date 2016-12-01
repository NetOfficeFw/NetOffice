using System;
using NetOffice;
namespace NetOffice.OWC10Api.Enums
{
	 /// <summary>
	 /// SupportByVersion OWC10 1
	 /// </summary>
	[SupportByVersionAttribute("OWC10", 1)]
	[EntityTypeAttribute(EntityType.IsEnum)]
	public enum XlReadingOrder
	{
		 /// <summary>
		 /// SupportByVersion OWC10 1
		 /// </summary>
		 /// <remarks>-5002</remarks>
		 [SupportByVersionAttribute("OWC10", 1)]
		 xlContext = -5002,

		 /// <summary>
		 /// SupportByVersion OWC10 1
		 /// </summary>
		 /// <remarks>-5003</remarks>
		 [SupportByVersionAttribute("OWC10", 1)]
		 xlLTR = -5003,

		 /// <summary>
		 /// SupportByVersion OWC10 1
		 /// </summary>
		 /// <remarks>-5004</remarks>
		 [SupportByVersionAttribute("OWC10", 1)]
		 xlRTL = -5004
	}
}