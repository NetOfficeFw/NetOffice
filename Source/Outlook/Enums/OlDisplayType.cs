using System;
using NetOffice;
namespace NetOffice.OutlookApi.Enums
{
	 /// <summary>
	 /// SupportByVersion Outlook 9, 10, 11, 12, 14
	 /// </summary>
	[SupportByVersionAttribute("Outlook", 9,10,11,12,14)]
	[EntityTypeAttribute(EntityType.IsEnum)]
	public enum OlDisplayType
	{
		 /// <summary>
		 /// SupportByVersion Outlook 9, 10, 11, 12, 14
		 /// </summary>
		 /// <remarks>0</remarks>
		 [SupportByVersionAttribute("Outlook", 9,10,11,12,14)]
		 olUser = 0,

		 /// <summary>
		 /// SupportByVersion Outlook 9, 10, 11, 12, 14
		 /// </summary>
		 /// <remarks>1</remarks>
		 [SupportByVersionAttribute("Outlook", 9,10,11,12,14)]
		 olDistList = 1,

		 /// <summary>
		 /// SupportByVersion Outlook 9, 10, 11, 12, 14
		 /// </summary>
		 /// <remarks>2</remarks>
		 [SupportByVersionAttribute("Outlook", 9,10,11,12,14)]
		 olForum = 2,

		 /// <summary>
		 /// SupportByVersion Outlook 9, 10, 11, 12, 14
		 /// </summary>
		 /// <remarks>3</remarks>
		 [SupportByVersionAttribute("Outlook", 9,10,11,12,14)]
		 olAgent = 3,

		 /// <summary>
		 /// SupportByVersion Outlook 9, 10, 11, 12, 14
		 /// </summary>
		 /// <remarks>4</remarks>
		 [SupportByVersionAttribute("Outlook", 9,10,11,12,14)]
		 olOrganization = 4,

		 /// <summary>
		 /// SupportByVersion Outlook 9, 10, 11, 12, 14
		 /// </summary>
		 /// <remarks>5</remarks>
		 [SupportByVersionAttribute("Outlook", 9,10,11,12,14)]
		 olPrivateDistList = 5,

		 /// <summary>
		 /// SupportByVersion Outlook 9, 10, 11, 12, 14
		 /// </summary>
		 /// <remarks>6</remarks>
		 [SupportByVersionAttribute("Outlook", 9,10,11,12,14)]
		 olRemoteUser = 6
	}
}