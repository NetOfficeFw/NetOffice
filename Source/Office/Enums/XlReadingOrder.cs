using System;
using NetOffice;
using NetOffice.Attributes;
namespace NetOffice.OfficeApi.Enums
{
	 /// <summary>
	 /// SupportByVersion Office 12, 14, 15, 16
	 /// </summary>
	[SupportByVersion("Office", 12,14,15,16)]
	[EntityType(EntityType.IsEnum)]
	public enum XlReadingOrder
	{
		 /// <summary>
		 /// SupportByVersion Office 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>-5002</remarks>
		 [SupportByVersion("Office", 12,14,15,16)]
		 xlContext = -5002,

		 /// <summary>
		 /// SupportByVersion Office 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>-5003</remarks>
		 [SupportByVersion("Office", 12,14,15,16)]
		 xlLTR = -5003,

		 /// <summary>
		 /// SupportByVersion Office 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>-5004</remarks>
		 [SupportByVersion("Office", 12,14,15,16)]
		 xlRTL = -5004
	}
}