using System;
using LateBindingApi.Core;
namespace NetOffice.OutlookApi.Enums
{
	 /// <summary>
	 /// SupportByVersion Outlook 11, 12, 14
	 /// </summary>
	[SupportByVersionAttribute("Outlook", 11,12,14)]
	[EntityTypeAttribute(EntityType.IsEnum)]
	public enum OlPermissionService
	{
		 /// <summary>
		 /// SupportByVersion Outlook 11, 12, 14
		 /// </summary>
		 /// <remarks>0</remarks>
		 [SupportByVersionAttribute("Outlook", 11,12,14)]
		 olUnknown = 0,

		 /// <summary>
		 /// SupportByVersion Outlook 11, 12, 14
		 /// </summary>
		 /// <remarks>1</remarks>
		 [SupportByVersionAttribute("Outlook", 11,12,14)]
		 olWindows = 1,

		 /// <summary>
		 /// SupportByVersion Outlook 11, 12, 14
		 /// </summary>
		 /// <remarks>2</remarks>
		 [SupportByVersionAttribute("Outlook", 11,12,14)]
		 olPassport = 2
	}
}