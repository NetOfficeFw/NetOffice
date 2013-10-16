using System;
using NetOffice;
namespace NetOffice.OutlookApi.Enums
{
	 /// <summary>
	 /// SupportByVersion Outlook 14, 15
	 /// </summary>
	[SupportByVersionAttribute("Outlook", 14,15)]
	[EntityTypeAttribute(EntityType.IsEnum)]
	public enum OlAppointmentCopyOptions
	{
		 /// <summary>
		 /// SupportByVersion Outlook 14, 15
		 /// </summary>
		 /// <remarks>0</remarks>
		 [SupportByVersionAttribute("Outlook", 14,15)]
		 olPromptUser = 0,

		 /// <summary>
		 /// SupportByVersion Outlook 14, 15
		 /// </summary>
		 /// <remarks>1</remarks>
		 [SupportByVersionAttribute("Outlook", 14,15)]
		 olCreateAppointment = 1,

		 /// <summary>
		 /// SupportByVersion Outlook 14, 15
		 /// </summary>
		 /// <remarks>2</remarks>
		 [SupportByVersionAttribute("Outlook", 14,15)]
		 olCopyAsAccept = 2
	}
}