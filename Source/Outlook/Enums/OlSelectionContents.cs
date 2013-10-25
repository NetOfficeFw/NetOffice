using System;
using NetOffice;
namespace NetOffice.OutlookApi.Enums
{
	 /// <summary>
	 /// SupportByVersion Outlook 14, 15
	 /// </summary>
	 ///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff868791.aspx </remarks>
	[SupportByVersionAttribute("Outlook", 14,15)]
	[EntityTypeAttribute(EntityType.IsEnum)]
	public enum OlSelectionContents
	{
		 /// <summary>
		 /// SupportByVersion Outlook 14, 15
		 /// </summary>
		 /// <remarks>1</remarks>
		 [SupportByVersionAttribute("Outlook", 14,15)]
		 olConversationHeaders = 1
	}
}