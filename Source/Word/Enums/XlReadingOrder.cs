using System;
using NetOffice;
namespace NetOffice.WordApi.Enums
{
	 /// <summary>
	 /// SupportByVersion Word 14, 15
	 /// </summary>
	 ///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff839947.aspx </remarks>
	[SupportByVersionAttribute("Word", 14,15)]
	[EntityTypeAttribute(EntityType.IsEnum)]
	public enum XlReadingOrder
	{
		 /// <summary>
		 /// SupportByVersion Word 14, 15
		 /// </summary>
		 /// <remarks>-5002</remarks>
		 [SupportByVersionAttribute("Word", 14,15)]
		 xlContext = -5002,

		 /// <summary>
		 /// SupportByVersion Word 14, 15
		 /// </summary>
		 /// <remarks>-5003</remarks>
		 [SupportByVersionAttribute("Word", 14,15)]
		 xlLTR = -5003,

		 /// <summary>
		 /// SupportByVersion Word 14, 15
		 /// </summary>
		 /// <remarks>-5004</remarks>
		 [SupportByVersionAttribute("Word", 14,15)]
		 xlRTL = -5004
	}
}