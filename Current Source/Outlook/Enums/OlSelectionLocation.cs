using System;
using LateBindingApi.Core;
namespace NetOffice.OutlookApi.Enums
{
	 /// <summary>
	 /// SupportByVersion Outlook 14
	 /// </summary>
	[SupportByVersionAttribute("Outlook", 14)]
	[EntityTypeAttribute(EntityType.IsEnum)]
	public enum OlSelectionLocation
	{
		 /// <summary>
		 /// SupportByVersion Outlook 14
		 /// </summary>
		 /// <remarks>0</remarks>
		 [SupportByVersionAttribute("Outlook", 14)]
		 olViewList = 0,

		 /// <summary>
		 /// SupportByVersion Outlook 14
		 /// </summary>
		 /// <remarks>1</remarks>
		 [SupportByVersionAttribute("Outlook", 14)]
		 olToDoBarTaskList = 1,

		 /// <summary>
		 /// SupportByVersion Outlook 14
		 /// </summary>
		 /// <remarks>2</remarks>
		 [SupportByVersionAttribute("Outlook", 14)]
		 olToDoBarAppointmentList = 2,

		 /// <summary>
		 /// SupportByVersion Outlook 14
		 /// </summary>
		 /// <remarks>3</remarks>
		 [SupportByVersionAttribute("Outlook", 14)]
		 olDailyTaskList = 3,

		 /// <summary>
		 /// SupportByVersion Outlook 14
		 /// </summary>
		 /// <remarks>4</remarks>
		 [SupportByVersionAttribute("Outlook", 14)]
		 olAttachmentWell = 4
	}
}