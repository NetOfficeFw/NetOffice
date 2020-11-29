using System;
using NetOffice;
using NetOffice.Attributes;
namespace NetOffice.WordApi.Enums
{
	 /// <summary>
	 /// SupportByVersion Word 14, 15, 16
	 /// </summary>
	 ///<remarks> MSDN Online Documentation: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.XlReadingOrder"/> </remarks>
	[SupportByVersion("Word", 14,15,16)]
	[EntityType(EntityType.IsEnum)]
	public enum XlReadingOrder
	{
		 /// <summary>
		 /// SupportByVersion Word 14, 15, 16
		 /// </summary>
		 /// <remarks>-5002</remarks>
		 [SupportByVersion("Word", 14,15,16)]
		 xlContext = -5002,

		 /// <summary>
		 /// SupportByVersion Word 14, 15, 16
		 /// </summary>
		 /// <remarks>-5003</remarks>
		 [SupportByVersion("Word", 14,15,16)]
		 xlLTR = -5003,

		 /// <summary>
		 /// SupportByVersion Word 14, 15, 16
		 /// </summary>
		 /// <remarks>-5004</remarks>
		 [SupportByVersion("Word", 14,15,16)]
		 xlRTL = -5004
	}
}