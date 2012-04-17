using System;
using LateBindingApi.Core;
namespace NetOffice.OutlookApi.Enums
{
	 /// <summary>
	 /// SupportByVersion Outlook 11, 12, 14
	 /// </summary>
	[SupportByVersionAttribute("Outlook", 11,12,14)]
	[EntityTypeAttribute(EntityType.IsEnum)]
	public enum OlExchangeConnectionMode
	{
		 /// <summary>
		 /// SupportByVersion Outlook 11, 12, 14
		 /// </summary>
		 /// <remarks>0</remarks>
		 [SupportByVersionAttribute("Outlook", 11,12,14)]
		 olNoExchange = 0,

		 /// <summary>
		 /// SupportByVersion Outlook 11, 12, 14
		 /// </summary>
		 /// <remarks>100</remarks>
		 [SupportByVersionAttribute("Outlook", 11,12,14)]
		 olOffline = 100,

		 /// <summary>
		 /// SupportByVersion Outlook 11, 12, 14
		 /// </summary>
		 /// <remarks>200</remarks>
		 [SupportByVersionAttribute("Outlook", 11,12,14)]
		 olCachedOffline = 200,

		 /// <summary>
		 /// SupportByVersion Outlook 11, 12, 14
		 /// </summary>
		 /// <remarks>300</remarks>
		 [SupportByVersionAttribute("Outlook", 11,12,14)]
		 olDisconnected = 300,

		 /// <summary>
		 /// SupportByVersion Outlook 11, 12, 14
		 /// </summary>
		 /// <remarks>400</remarks>
		 [SupportByVersionAttribute("Outlook", 11,12,14)]
		 olCachedDisconnected = 400,

		 /// <summary>
		 /// SupportByVersion Outlook 11, 12, 14
		 /// </summary>
		 /// <remarks>500</remarks>
		 [SupportByVersionAttribute("Outlook", 11,12,14)]
		 olCachedConnectedHeaders = 500,

		 /// <summary>
		 /// SupportByVersion Outlook 11, 12, 14
		 /// </summary>
		 /// <remarks>600</remarks>
		 [SupportByVersionAttribute("Outlook", 11,12,14)]
		 olCachedConnectedDrizzle = 600,

		 /// <summary>
		 /// SupportByVersion Outlook 11, 12, 14
		 /// </summary>
		 /// <remarks>700</remarks>
		 [SupportByVersionAttribute("Outlook", 11,12,14)]
		 olCachedConnectedFull = 700,

		 /// <summary>
		 /// SupportByVersion Outlook 11, 12, 14
		 /// </summary>
		 /// <remarks>800</remarks>
		 [SupportByVersionAttribute("Outlook", 11,12,14)]
		 olOnline = 800
	}
}