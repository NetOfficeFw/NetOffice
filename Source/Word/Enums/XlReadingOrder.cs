using System;
using NetOffice;
namespace NetOffice.WordApi.Enums
{
	 /// <summary>
	 /// SupportByVersion Word 14
	 /// </summary>
	[SupportByVersionAttribute("Word", 14)]
	[EntityTypeAttribute(EntityType.IsEnum)]
	public enum XlReadingOrder
	{
		 /// <summary>
		 /// SupportByVersion Word 14
		 /// </summary>
		 /// <remarks>-5002</remarks>
		 [SupportByVersionAttribute("Word", 14)]
		 xlContext = -5002,

		 /// <summary>
		 /// SupportByVersion Word 14
		 /// </summary>
		 /// <remarks>-5003</remarks>
		 [SupportByVersionAttribute("Word", 14)]
		 xlLTR = -5003,

		 /// <summary>
		 /// SupportByVersion Word 14
		 /// </summary>
		 /// <remarks>-5004</remarks>
		 [SupportByVersionAttribute("Word", 14)]
		 xlRTL = -5004
	}
}