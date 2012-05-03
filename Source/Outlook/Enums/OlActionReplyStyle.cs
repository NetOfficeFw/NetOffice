using System;
using NetOffice;
namespace NetOffice.OutlookApi.Enums
{
	 /// <summary>
	 /// SupportByVersion Outlook 9, 10, 11, 12, 14
	 /// </summary>
	[SupportByVersionAttribute("Outlook", 9,10,11,12,14)]
	[EntityTypeAttribute(EntityType.IsEnum)]
	public enum OlActionReplyStyle
	{
		 /// <summary>
		 /// SupportByVersion Outlook 9, 10, 11, 12, 14
		 /// </summary>
		 /// <remarks>0</remarks>
		 [SupportByVersionAttribute("Outlook", 9,10,11,12,14)]
		 olOmitOriginalText = 0,

		 /// <summary>
		 /// SupportByVersion Outlook 9, 10, 11, 12, 14
		 /// </summary>
		 /// <remarks>1</remarks>
		 [SupportByVersionAttribute("Outlook", 9,10,11,12,14)]
		 olEmbedOriginalItem = 1,

		 /// <summary>
		 /// SupportByVersion Outlook 9, 10, 11, 12, 14
		 /// </summary>
		 /// <remarks>2</remarks>
		 [SupportByVersionAttribute("Outlook", 9,10,11,12,14)]
		 olIncludeOriginalText = 2,

		 /// <summary>
		 /// SupportByVersion Outlook 9, 10, 11, 12, 14
		 /// </summary>
		 /// <remarks>3</remarks>
		 [SupportByVersionAttribute("Outlook", 9,10,11,12,14)]
		 olIndentOriginalText = 3,

		 /// <summary>
		 /// SupportByVersion Outlook 9, 10, 11, 12, 14
		 /// </summary>
		 /// <remarks>4</remarks>
		 [SupportByVersionAttribute("Outlook", 9,10,11,12,14)]
		 olLinkOriginalItem = 4,

		 /// <summary>
		 /// SupportByVersion Outlook 9, 10, 11, 12, 14
		 /// </summary>
		 /// <remarks>5</remarks>
		 [SupportByVersionAttribute("Outlook", 9,10,11,12,14)]
		 olUserPreference = 5,

		 /// <summary>
		 /// SupportByVersion Outlook 9, 10, 11, 12, 14
		 /// </summary>
		 /// <remarks>1000</remarks>
		 [SupportByVersionAttribute("Outlook", 9,10,11,12,14)]
		 olReplyTickOriginalText = 1000
	}
}