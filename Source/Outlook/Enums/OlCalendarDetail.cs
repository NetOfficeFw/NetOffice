using System;
using NetOffice;
namespace NetOffice.OutlookApi.Enums
{
	 /// <summary>
	 /// SupportByVersion Outlook 12, 14
	 /// </summary>
	[SupportByVersionAttribute("Outlook", 12,14)]
	[EntityTypeAttribute(EntityType.IsEnum)]
	public enum OlCalendarDetail
	{
		 /// <summary>
		 /// SupportByVersion Outlook 12, 14
		 /// </summary>
		 /// <remarks>0</remarks>
		 [SupportByVersionAttribute("Outlook", 12,14)]
		 olFreeBusyOnly = 0,

		 /// <summary>
		 /// SupportByVersion Outlook 12, 14
		 /// </summary>
		 /// <remarks>1</remarks>
		 [SupportByVersionAttribute("Outlook", 12,14)]
		 olFreeBusyAndSubject = 1,

		 /// <summary>
		 /// SupportByVersion Outlook 12, 14
		 /// </summary>
		 /// <remarks>2</remarks>
		 [SupportByVersionAttribute("Outlook", 12,14)]
		 olFullDetails = 2
	}
}