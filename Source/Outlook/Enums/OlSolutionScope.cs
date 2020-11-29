using System;
using NetOffice;
using NetOffice.Attributes;
namespace NetOffice.OutlookApi.Enums
{
	 /// <summary>
	 /// SupportByVersion Outlook 14, 15, 16
	 /// </summary>
	 ///<remarks> MSDN Online Documentation: <see href="https://docs.microsoft.com/en-us/office/vba/api/Outlook.OlSolutionScope"/> </remarks>
	[SupportByVersion("Outlook", 14,15,16)]
	[EntityType(EntityType.IsEnum)]
	public enum OlSolutionScope
	{
		 /// <summary>
		 /// SupportByVersion Outlook 14, 15, 16
		 /// </summary>
		 /// <remarks>0</remarks>
		 [SupportByVersion("Outlook", 14,15,16)]
		 olHideInDefaultModules = 0,

		 /// <summary>
		 /// SupportByVersion Outlook 14, 15, 16
		 /// </summary>
		 /// <remarks>1</remarks>
		 [SupportByVersion("Outlook", 14,15,16)]
		 olShowInDefaultModules = 1
	}
}