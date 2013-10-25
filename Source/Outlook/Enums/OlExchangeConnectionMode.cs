using System;
using NetOffice;
namespace NetOffice.OutlookApi.Enums
{
	 /// <summary>
	 /// SupportByVersion Outlook 11, 12, 14, 15
	 /// </summary>
	 ///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff868474.aspx </remarks>
	[SupportByVersionAttribute("Outlook", 11,12,14,15)]
	[EntityTypeAttribute(EntityType.IsEnum)]
	public enum OlExchangeConnectionMode
	{
		 /// <summary>
		 /// SupportByVersion Outlook 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>0</remarks>
		 [SupportByVersionAttribute("Outlook", 11,12,14,15)]
		 olNoExchange = 0,

		 /// <summary>
		 /// SupportByVersion Outlook 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>100</remarks>
		 [SupportByVersionAttribute("Outlook", 11,12,14,15)]
		 olOffline = 100,

		 /// <summary>
		 /// SupportByVersion Outlook 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>200</remarks>
		 [SupportByVersionAttribute("Outlook", 11,12,14,15)]
		 olCachedOffline = 200,

		 /// <summary>
		 /// SupportByVersion Outlook 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>300</remarks>
		 [SupportByVersionAttribute("Outlook", 11,12,14,15)]
		 olDisconnected = 300,

		 /// <summary>
		 /// SupportByVersion Outlook 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>400</remarks>
		 [SupportByVersionAttribute("Outlook", 11,12,14,15)]
		 olCachedDisconnected = 400,

		 /// <summary>
		 /// SupportByVersion Outlook 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>500</remarks>
		 [SupportByVersionAttribute("Outlook", 11,12,14,15)]
		 olCachedConnectedHeaders = 500,

		 /// <summary>
		 /// SupportByVersion Outlook 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>600</remarks>
		 [SupportByVersionAttribute("Outlook", 11,12,14,15)]
		 olCachedConnectedDrizzle = 600,

		 /// <summary>
		 /// SupportByVersion Outlook 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>700</remarks>
		 [SupportByVersionAttribute("Outlook", 11,12,14,15)]
		 olCachedConnectedFull = 700,

		 /// <summary>
		 /// SupportByVersion Outlook 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>800</remarks>
		 [SupportByVersionAttribute("Outlook", 11,12,14,15)]
		 olOnline = 800
	}
}