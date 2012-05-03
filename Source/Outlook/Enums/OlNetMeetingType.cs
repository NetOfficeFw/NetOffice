using System;
using NetOffice;
namespace NetOffice.OutlookApi.Enums
{
	 /// <summary>
	 /// SupportByVersion Outlook 9, 10, 11, 12, 14
	 /// </summary>
	[SupportByVersionAttribute("Outlook", 9,10,11,12,14)]
	[EntityTypeAttribute(EntityType.IsEnum)]
	public enum OlNetMeetingType
	{
		 /// <summary>
		 /// SupportByVersion Outlook 9, 10, 11, 12, 14
		 /// </summary>
		 /// <remarks>0</remarks>
		 [SupportByVersionAttribute("Outlook", 9,10,11,12,14)]
		 olNetMeeting = 0,

		 /// <summary>
		 /// SupportByVersion Outlook 9, 10, 11, 12, 14
		 /// </summary>
		 /// <remarks>1</remarks>
		 [SupportByVersionAttribute("Outlook", 9,10,11,12,14)]
		 olNetShow = 1,

		 /// <summary>
		 /// SupportByVersion Outlook 9
		 /// </summary>
		 /// <remarks>2</remarks>
		 [SupportByVersionAttribute("Outlook", 9)]
		 olChat = 2,

		 /// <summary>
		 /// SupportByVersion Outlook 10, 11, 12, 14
		 /// </summary>
		 /// <remarks>2</remarks>
		 [SupportByVersionAttribute("Outlook", 10,11,12,14)]
		 olExchangeConferencing = 2
	}
}