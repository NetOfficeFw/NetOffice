using System;
using NetOffice;
using NetOffice.Attributes;
namespace NetOffice.OutlookApi.Enums
{
	 /// <summary>
	 /// SupportByVersion Outlook 11, 12, 14, 15, 16
	 /// </summary>
	 ///<remarks> MSDN Online Documentation: <see href="https://docs.microsoft.com/en-us/office/vba/api/Outlook.OlPermissionService"/> </remarks>
	[SupportByVersion("Outlook", 11,12,14,15,16)]
	[EntityType(EntityType.IsEnum)]
	public enum OlPermissionService
	{
		 /// <summary>
		 /// SupportByVersion Outlook 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>0</remarks>
		 [SupportByVersion("Outlook", 11,12,14,15,16)]
		 olUnknown = 0,

		 /// <summary>
		 /// SupportByVersion Outlook 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>1</remarks>
		 [SupportByVersion("Outlook", 11,12,14,15,16)]
		 olWindows = 1,

		 /// <summary>
		 /// SupportByVersion Outlook 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>2</remarks>
		 [SupportByVersion("Outlook", 11,12,14,15,16)]
		 olPassport = 2
	}
}