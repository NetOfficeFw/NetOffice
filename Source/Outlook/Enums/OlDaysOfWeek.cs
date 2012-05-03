using System;
using NetOffice;
namespace NetOffice.OutlookApi.Enums
{
	 /// <summary>
	 /// SupportByVersion Outlook 9, 10, 11, 12, 14
	 /// </summary>
	[SupportByVersionAttribute("Outlook", 9,10,11,12,14)]
	[EntityTypeAttribute(EntityType.IsEnum)]
	public enum OlDaysOfWeek
	{
		 /// <summary>
		 /// SupportByVersion Outlook 9, 10, 11, 12, 14
		 /// </summary>
		 /// <remarks>1</remarks>
		 [SupportByVersionAttribute("Outlook", 9,10,11,12,14)]
		 olSunday = 1,

		 /// <summary>
		 /// SupportByVersion Outlook 9, 10, 11, 12, 14
		 /// </summary>
		 /// <remarks>2</remarks>
		 [SupportByVersionAttribute("Outlook", 9,10,11,12,14)]
		 olMonday = 2,

		 /// <summary>
		 /// SupportByVersion Outlook 9, 10, 11, 12, 14
		 /// </summary>
		 /// <remarks>4</remarks>
		 [SupportByVersionAttribute("Outlook", 9,10,11,12,14)]
		 olTuesday = 4,

		 /// <summary>
		 /// SupportByVersion Outlook 9, 10, 11, 12, 14
		 /// </summary>
		 /// <remarks>8</remarks>
		 [SupportByVersionAttribute("Outlook", 9,10,11,12,14)]
		 olWednesday = 8,

		 /// <summary>
		 /// SupportByVersion Outlook 9, 10, 11, 12, 14
		 /// </summary>
		 /// <remarks>16</remarks>
		 [SupportByVersionAttribute("Outlook", 9,10,11,12,14)]
		 olThursday = 16,

		 /// <summary>
		 /// SupportByVersion Outlook 9, 10, 11, 12, 14
		 /// </summary>
		 /// <remarks>32</remarks>
		 [SupportByVersionAttribute("Outlook", 9,10,11,12,14)]
		 olFriday = 32,

		 /// <summary>
		 /// SupportByVersion Outlook 9, 10, 11, 12, 14
		 /// </summary>
		 /// <remarks>64</remarks>
		 [SupportByVersionAttribute("Outlook", 9,10,11,12,14)]
		 olSaturday = 64
	}
}