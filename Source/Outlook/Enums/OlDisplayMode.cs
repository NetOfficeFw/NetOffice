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
	public enum OlDisplayMode
	{
		 /// <summary>
		 /// SupportByVersion Outlook 16
		 /// </summary>
		 /// <remarks>0</remarks>
		 [SupportByVersion("Outlook", 16)]
		 olDisplayModeNormal = 0,

		 /// <summary>
		 /// SupportByVersion Outlook 16
		 /// </summary>
		 /// <remarks>1</remarks>
		 [SupportByVersion("Outlook", 16)]
		 olDisplayModePortraitView = 1,

		 /// <summary>
		 /// SupportByVersion Outlook 16
		 /// </summary>
		 /// <remarks>2</remarks>
		 [SupportByVersion("Outlook", 16)]
		 olDisplayModePortraitReadingPane = 2
	}
}