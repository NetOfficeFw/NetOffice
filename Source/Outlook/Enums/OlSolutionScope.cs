using System;
using NetOffice;
namespace NetOffice.OutlookApi.Enums
{
	 /// <summary>
	 /// SupportByVersion Outlook 14
	 /// </summary>
	[SupportByVersionAttribute("Outlook", 14)]
	[EntityTypeAttribute(EntityType.IsEnum)]
	public enum OlSolutionScope
	{
		 /// <summary>
		 /// SupportByVersion Outlook 14
		 /// </summary>
		 /// <remarks>0</remarks>
		 [SupportByVersionAttribute("Outlook", 14)]
		 olHideInDefaultModules = 0,

		 /// <summary>
		 /// SupportByVersion Outlook 14
		 /// </summary>
		 /// <remarks>1</remarks>
		 [SupportByVersionAttribute("Outlook", 14)]
		 olShowInDefaultModules = 1
	}
}