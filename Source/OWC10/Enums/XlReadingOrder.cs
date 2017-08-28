using System;
using NetOffice;
using NetOffice.Attributes;
namespace NetOffice.OWC10Api.Enums
{
	 /// <summary>
	 /// SupportByVersion OWC10 1
	 /// </summary>
	[SupportByVersion("OWC10", 1)]
	[EntityType(EntityType.IsEnum)]
	public enum XlReadingOrder
	{
		 /// <summary>
		 /// SupportByVersion OWC10 1
		 /// </summary>
		 /// <remarks>-5002</remarks>
		 [SupportByVersion("OWC10", 1)]
		 xlContext = -5002,

		 /// <summary>
		 /// SupportByVersion OWC10 1
		 /// </summary>
		 /// <remarks>-5003</remarks>
		 [SupportByVersion("OWC10", 1)]
		 xlLTR = -5003,

		 /// <summary>
		 /// SupportByVersion OWC10 1
		 /// </summary>
		 /// <remarks>-5004</remarks>
		 [SupportByVersion("OWC10", 1)]
		 xlRTL = -5004
	}
}