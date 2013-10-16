using System;
using NetOffice;
namespace NetOffice.OfficeApi.Enums
{
	 /// <summary>
	 /// SupportByVersion Office 12, 14, 15
	 /// </summary>
	[SupportByVersionAttribute("Office", 12,14,15)]
	[EntityTypeAttribute(EntityType.IsEnum)]
	public enum XlReadingOrder
	{
		 /// <summary>
		 /// SupportByVersion Office 12, 14, 15
		 /// </summary>
		 /// <remarks>-5002</remarks>
		 [SupportByVersionAttribute("Office", 12,14,15)]
		 xlContext = -5002,

		 /// <summary>
		 /// SupportByVersion Office 12, 14, 15
		 /// </summary>
		 /// <remarks>-5003</remarks>
		 [SupportByVersionAttribute("Office", 12,14,15)]
		 xlLTR = -5003,

		 /// <summary>
		 /// SupportByVersion Office 12, 14, 15
		 /// </summary>
		 /// <remarks>-5004</remarks>
		 [SupportByVersionAttribute("Office", 12,14,15)]
		 xlRTL = -5004
	}
}