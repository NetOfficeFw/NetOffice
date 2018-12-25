using System;
using NetOffice;
using NetOffice.Attributes;
namespace NetOffice.OutlookApi.Enums
{
	 /// <summary>
	 /// SupportByVersion Outlook 16
	 /// </summary>
	[SupportByVersion("Outlook", 16)]
	[EntityType(EntityType.IsEnum)]
	public enum OlUnifiedGroupFolderType
	{
		 /// <summary>
		 /// SupportByVersion Outlook 16
		 /// </summary>
		 /// <remarks>0</remarks>
		 [SupportByVersion("Outlook", 16)]
		 olGroupMailFolder = 0,

		 /// <summary>
		 /// SupportByVersion Outlook 16
		 /// </summary>
		 /// <remarks>1</remarks>
		 [SupportByVersion("Outlook", 16)]
		 olGroupCalendarFolder = 1
	}
}