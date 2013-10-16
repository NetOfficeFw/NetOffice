using System;
using NetOffice;
namespace NetOffice.OutlookApi.Enums
{
	 /// <summary>
	 /// SupportByVersion Outlook 10, 11, 12, 14, 15
	 /// </summary>
	[SupportByVersionAttribute("Outlook", 10,11,12,14,15)]
	[EntityTypeAttribute(EntityType.IsEnum)]
	public enum OlDownloadState
	{
		 /// <summary>
		 /// SupportByVersion Outlook 10, 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>0</remarks>
		 [SupportByVersionAttribute("Outlook", 10,11,12,14,15)]
		 olHeaderOnly = 0,

		 /// <summary>
		 /// SupportByVersion Outlook 10, 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>1</remarks>
		 [SupportByVersionAttribute("Outlook", 10,11,12,14,15)]
		 olFullItem = 1
	}
}