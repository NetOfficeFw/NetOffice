using System;
using LateBindingApi.Core;
namespace NetOffice.OutlookApi.Enums
{
	 /// <summary>
	 /// SupportByVersion Outlook 14
	 /// </summary>
	[SupportByVersionAttribute("Outlook", 14)]
	[EntityTypeAttribute(EntityType.IsEnum)]
	public enum OlMobileFormat
	{
		 /// <summary>
		 /// SupportByVersion Outlook 14
		 /// </summary>
		 /// <remarks>0</remarks>
		 [SupportByVersionAttribute("Outlook", 14)]
		 olSMS = 0,

		 /// <summary>
		 /// SupportByVersion Outlook 14
		 /// </summary>
		 /// <remarks>1</remarks>
		 [SupportByVersionAttribute("Outlook", 14)]
		 olMMS = 1
	}
}