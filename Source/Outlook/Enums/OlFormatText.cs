using System;
using NetOffice;
using NetOffice.Attributes;
namespace NetOffice.OutlookApi.Enums
{
	 /// <summary>
	 /// SupportByVersion Outlook 12, 14, 15, 16
	 /// </summary>
	 ///<remarks> MSDN Online Documentation: <see href="https://docs.microsoft.com/en-us/office/vba/api/Outlook.OlFormatText"/> </remarks>
	[SupportByVersion("Outlook", 12,14,15,16)]
	[EntityType(EntityType.IsEnum)]
	public enum OlFormatText
	{
		 /// <summary>
		 /// SupportByVersion Outlook 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>1</remarks>
		 [SupportByVersion("Outlook", 12,14,15,16)]
		 olFormatTextText = 1
	}
}