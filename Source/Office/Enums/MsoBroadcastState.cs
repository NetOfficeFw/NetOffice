using System;
using NetOffice;
using NetOffice.Attributes;
namespace NetOffice.OfficeApi.Enums
{
	 /// <summary>
	 /// SupportByVersion Office 15,16
	 /// </summary>
	 ///<remarks> MSDN Online Documentation: <see href="https://docs.microsoft.com/en-us/office/vba/api/overview/Library-Reference/msobroadcaststate-enumeration-office"/> </remarks>
	[SupportByVersion("Office", 15, 16)]
	[EntityType(EntityType.IsEnum)]
	public enum MsoBroadcastState
	{
		 /// <summary>
		 /// SupportByVersion Office 15,16
		 /// </summary>
		 /// <remarks>0</remarks>
		 [SupportByVersion("Office", 15, 16)]
		 NoBroadcast = 0,

		 /// <summary>
		 /// SupportByVersion Office 15,16
		 /// </summary>
		 /// <remarks>1</remarks>
		 [SupportByVersion("Office", 15, 16)]
		 BroadcastStarted = 1,

		 /// <summary>
		 /// SupportByVersion Office 15,16
		 /// </summary>
		 /// <remarks>2</remarks>
		 [SupportByVersion("Office", 15, 16)]
		 BroadcastPaused = 2
	}
}