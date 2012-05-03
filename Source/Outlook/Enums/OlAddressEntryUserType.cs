using System;
using NetOffice;
namespace NetOffice.OutlookApi.Enums
{
	 /// <summary>
	 /// SupportByVersion Outlook 12, 14
	 /// </summary>
	[SupportByVersionAttribute("Outlook", 12,14)]
	[EntityTypeAttribute(EntityType.IsEnum)]
	public enum OlAddressEntryUserType
	{
		 /// <summary>
		 /// SupportByVersion Outlook 12, 14
		 /// </summary>
		 /// <remarks>0</remarks>
		 [SupportByVersionAttribute("Outlook", 12,14)]
		 olExchangeUserAddressEntry = 0,

		 /// <summary>
		 /// SupportByVersion Outlook 12, 14
		 /// </summary>
		 /// <remarks>1</remarks>
		 [SupportByVersionAttribute("Outlook", 12,14)]
		 olExchangeDistributionListAddressEntry = 1,

		 /// <summary>
		 /// SupportByVersion Outlook 12, 14
		 /// </summary>
		 /// <remarks>2</remarks>
		 [SupportByVersionAttribute("Outlook", 12,14)]
		 olExchangePublicFolderAddressEntry = 2,

		 /// <summary>
		 /// SupportByVersion Outlook 12, 14
		 /// </summary>
		 /// <remarks>3</remarks>
		 [SupportByVersionAttribute("Outlook", 12,14)]
		 olExchangeAgentAddressEntry = 3,

		 /// <summary>
		 /// SupportByVersion Outlook 12, 14
		 /// </summary>
		 /// <remarks>4</remarks>
		 [SupportByVersionAttribute("Outlook", 12,14)]
		 olExchangeOrganizationAddressEntry = 4,

		 /// <summary>
		 /// SupportByVersion Outlook 12, 14
		 /// </summary>
		 /// <remarks>5</remarks>
		 [SupportByVersionAttribute("Outlook", 12,14)]
		 olExchangeRemoteUserAddressEntry = 5,

		 /// <summary>
		 /// SupportByVersion Outlook 12, 14
		 /// </summary>
		 /// <remarks>10</remarks>
		 [SupportByVersionAttribute("Outlook", 12,14)]
		 olOutlookContactAddressEntry = 10,

		 /// <summary>
		 /// SupportByVersion Outlook 12, 14
		 /// </summary>
		 /// <remarks>11</remarks>
		 [SupportByVersionAttribute("Outlook", 12,14)]
		 olOutlookDistributionListAddressEntry = 11,

		 /// <summary>
		 /// SupportByVersion Outlook 12, 14
		 /// </summary>
		 /// <remarks>20</remarks>
		 [SupportByVersionAttribute("Outlook", 12,14)]
		 olLdapAddressEntry = 20,

		 /// <summary>
		 /// SupportByVersion Outlook 12, 14
		 /// </summary>
		 /// <remarks>30</remarks>
		 [SupportByVersionAttribute("Outlook", 12,14)]
		 olSmtpAddressEntry = 30,

		 /// <summary>
		 /// SupportByVersion Outlook 12, 14
		 /// </summary>
		 /// <remarks>40</remarks>
		 [SupportByVersionAttribute("Outlook", 12,14)]
		 olOtherAddressEntry = 40
	}
}