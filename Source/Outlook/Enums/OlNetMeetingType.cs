using System;
using NetOffice;
using NetOffice.Attributes;
namespace NetOffice.OutlookApi.Enums
{
	 /// <summary>
	 /// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
	 /// </summary>
	 ///<remarks> MSDN Online Documentation: <see href="https://docs.microsoft.com/en-us/"/> </remarks>
	[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
	[EntityType(EntityType.IsEnum)]
	public enum OlNetMeetingType
	{
		 /// <summary>
		 /// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>0</remarks>
		 [SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		 olNetMeeting = 0,

		 /// <summary>
		 /// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>1</remarks>
		 [SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		 olNetShow = 1,

		 /// <summary>
		 /// SupportByVersion Outlook 9
		 /// </summary>
		 /// <remarks>2</remarks>
		 [SupportByVersion("Outlook", 9)]
		 olChat = 2,

		 /// <summary>
		 /// SupportByVersion Outlook 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>2</remarks>
		 [SupportByVersion("Outlook", 10,11,12,14,15,16)]
		 olExchangeConferencing = 2
	}
}