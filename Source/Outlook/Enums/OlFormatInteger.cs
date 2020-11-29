using System;
using NetOffice;
using NetOffice.Attributes;
namespace NetOffice.OutlookApi.Enums
{
	 /// <summary>
	 /// SupportByVersion Outlook 12, 14, 15, 16
	 /// </summary>
	 ///<remarks> MSDN Online Documentation: <see href="https://docs.microsoft.com/en-us/office/vba/api/Outlook.OlFormatInteger"/> </remarks>
	[SupportByVersion("Outlook", 12,14,15,16)]
	[EntityType(EntityType.IsEnum)]
	public enum OlFormatInteger
	{
		 /// <summary>
		 /// SupportByVersion Outlook 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>1</remarks>
		 [SupportByVersion("Outlook", 12,14,15,16)]
		 olFormatIntegerPlain = 1,

		 /// <summary>
		 /// SupportByVersion Outlook 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>2</remarks>
		 [SupportByVersion("Outlook", 12,14,15,16)]
		 olFormatIntegerComputer1 = 2,

		 /// <summary>
		 /// SupportByVersion Outlook 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>3</remarks>
		 [SupportByVersion("Outlook", 12,14,15,16)]
		 olFormatIntegerComputer2 = 3,

		 /// <summary>
		 /// SupportByVersion Outlook 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>4</remarks>
		 [SupportByVersion("Outlook", 12,14,15,16)]
		 olFormatIntegerComputer3 = 4
	}
}