using System;
using NetOffice;
using NetOffice.Attributes;
namespace NetOffice.OutlookApi.Enums
{
	 /// <summary>
	 /// SupportByVersion Outlook 14, 15, 16
	 /// </summary>
	 ///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff868791.aspx </remarks>
	[SupportByVersion("Outlook", 14,15,16)]
	[EntityType(EntityType.IsEnum)]
	public enum OlSelectionContents
	{
		 /// <summary>
		 /// SupportByVersion Outlook 14, 15, 16
		 /// </summary>
		 /// <remarks>1</remarks>
		 [SupportByVersion("Outlook", 14,15,16)]
		 olConversationHeaders = 1
	}
}