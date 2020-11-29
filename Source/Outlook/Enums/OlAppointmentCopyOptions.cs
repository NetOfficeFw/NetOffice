using System;
using NetOffice;
using NetOffice.Attributes;
namespace NetOffice.OutlookApi.Enums
{
	 /// <summary>
	 /// SupportByVersion Outlook 14, 15, 16
	 /// </summary>
	 ///<remarks> MSDN Online Documentation: <see href="https://docs.microsoft.com/en-us/office/vba/api/Outlook.OlAppointmentCopyOptions"/> </remarks>
	[SupportByVersion("Outlook", 14,15,16)]
	[EntityType(EntityType.IsEnum)]
	public enum OlAppointmentCopyOptions
	{
		 /// <summary>
		 /// SupportByVersion Outlook 14, 15, 16
		 /// </summary>
		 /// <remarks>0</remarks>
		 [SupportByVersion("Outlook", 14,15,16)]
		 olPromptUser = 0,

		 /// <summary>
		 /// SupportByVersion Outlook 14, 15, 16
		 /// </summary>
		 /// <remarks>1</remarks>
		 [SupportByVersion("Outlook", 14,15,16)]
		 olCreateAppointment = 1,

		 /// <summary>
		 /// SupportByVersion Outlook 14, 15, 16
		 /// </summary>
		 /// <remarks>2</remarks>
		 [SupportByVersion("Outlook", 14,15,16)]
		 olCopyAsAccept = 2
	}
}