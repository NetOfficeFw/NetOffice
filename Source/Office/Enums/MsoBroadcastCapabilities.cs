using System;
using NetOffice;
using NetOffice.Attributes;
namespace NetOffice.OfficeApi.Enums
{
	 /// <summary>
	 /// SupportByVersion Office 15,16
	 /// </summary>
	 ///<remarks> MSDN Online Documentation: <see href="https://docs.microsoft.com/en-us/office/vba/api/overview/Library-Reference/msobroadcastcapabilities-enumeration-office"/> </remarks>
	[SupportByVersion("Office", 15, 16)]
	[EntityType(EntityType.IsEnum)]
	public enum MsoBroadcastCapabilities
	{
		 /// <summary>
		 /// SupportByVersion Office 15,16
		 /// </summary>
		 /// <remarks>1</remarks>
		 [SupportByVersion("Office", 15, 16)]
		 BroadcastCapFileSizeLimited = 1,

		 /// <summary>
		 /// SupportByVersion Office 15,16
		 /// </summary>
		 /// <remarks>2</remarks>
		 [SupportByVersion("Office", 15, 16)]
		 BroadcastCapSupportsMeetingNotes = 2,

		 /// <summary>
		 /// SupportByVersion Office 15,16
		 /// </summary>
		 /// <remarks>4</remarks>
		 [SupportByVersion("Office", 15, 16)]
		 BroadcastCapSupportsUpdateDoc = 4
	}
}