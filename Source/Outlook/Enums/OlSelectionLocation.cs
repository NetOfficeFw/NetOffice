using System;
using NetOffice;
using NetOffice.Attributes;
namespace NetOffice.OutlookApi.Enums
{
	 /// <summary>
	 /// SupportByVersion Outlook 14, 15, 16
	 /// </summary>
	 ///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff868261.aspx </remarks>
	[SupportByVersion("Outlook", 14,15,16)]
	[EntityType(EntityType.IsEnum)]
	public enum OlSelectionLocation
	{
		 /// <summary>
		 /// SupportByVersion Outlook 14, 15, 16
		 /// </summary>
		 /// <remarks>0</remarks>
		 [SupportByVersion("Outlook", 14,15,16)]
		 olViewList = 0,

		 /// <summary>
		 /// SupportByVersion Outlook 14, 15, 16
		 /// </summary>
		 /// <remarks>1</remarks>
		 [SupportByVersion("Outlook", 14,15,16)]
		 olToDoBarTaskList = 1,

		 /// <summary>
		 /// SupportByVersion Outlook 14, 15, 16
		 /// </summary>
		 /// <remarks>2</remarks>
		 [SupportByVersion("Outlook", 14,15,16)]
		 olToDoBarAppointmentList = 2,

		 /// <summary>
		 /// SupportByVersion Outlook 14, 15, 16
		 /// </summary>
		 /// <remarks>3</remarks>
		 [SupportByVersion("Outlook", 14,15,16)]
		 olDailyTaskList = 3,

		 /// <summary>
		 /// SupportByVersion Outlook 14, 15, 16
		 /// </summary>
		 /// <remarks>4</remarks>
		 [SupportByVersion("Outlook", 14,15,16)]
		 olAttachmentWell = 4
	}
}