using System;
using NetOffice;
using NetOffice.Attributes;
namespace NetOffice.OutlookApi.Enums
{
	 /// <summary>
	 /// SupportByVersion Outlook 10, 11, 12, 14, 15, 16
	 /// </summary>
	 ///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff869904.aspx </remarks>
	[SupportByVersion("Outlook", 10,11,12,14,15,16)]
	[EntityType(EntityType.IsEnum)]
	public enum OlViewType
	{
		 /// <summary>
		 /// SupportByVersion Outlook 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>0</remarks>
		 [SupportByVersion("Outlook", 10,11,12,14,15,16)]
		 olTableView = 0,

		 /// <summary>
		 /// SupportByVersion Outlook 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>1</remarks>
		 [SupportByVersion("Outlook", 10,11,12,14,15,16)]
		 olCardView = 1,

		 /// <summary>
		 /// SupportByVersion Outlook 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>2</remarks>
		 [SupportByVersion("Outlook", 10,11,12,14,15,16)]
		 olCalendarView = 2,

		 /// <summary>
		 /// SupportByVersion Outlook 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>3</remarks>
		 [SupportByVersion("Outlook", 10,11,12,14,15,16)]
		 olIconView = 3,

		 /// <summary>
		 /// SupportByVersion Outlook 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>4</remarks>
		 [SupportByVersion("Outlook", 10,11,12,14,15,16)]
		 olTimelineView = 4,

		 /// <summary>
		 /// SupportByVersion Outlook 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>5</remarks>
		 [SupportByVersion("Outlook", 12,14,15,16)]
		 olBusinessCardView = 5,

		 /// <summary>
		 /// SupportByVersion Outlook 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>6</remarks>
		 [SupportByVersion("Outlook", 12,14,15,16)]
		 olDailyTaskListView = 6,

		 /// <summary>
		 /// SupportByVersion Outlook 15,16
		 /// </summary>
		 /// <remarks>7</remarks>
		 [SupportByVersion("Outlook", 15, 16)]
		 olPeopleView = 7
	}
}