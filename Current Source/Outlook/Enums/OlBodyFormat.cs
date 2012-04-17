using System;
using LateBindingApi.Core;
namespace NetOffice.OutlookApi.Enums
{
	 /// <summary>
	 /// SupportByVersion Outlook 10, 11, 12, 14
	 /// </summary>
	[SupportByVersionAttribute("Outlook", 10,11,12,14)]
	[EntityTypeAttribute(EntityType.IsEnum)]
	public enum OlBodyFormat
	{
		 /// <summary>
		 /// SupportByVersion Outlook 10, 11, 12, 14
		 /// </summary>
		 /// <remarks>0</remarks>
		 [SupportByVersionAttribute("Outlook", 10,11,12,14)]
		 olFormatUnspecified = 0,

		 /// <summary>
		 /// SupportByVersion Outlook 10, 11, 12, 14
		 /// </summary>
		 /// <remarks>1</remarks>
		 [SupportByVersionAttribute("Outlook", 10,11,12,14)]
		 olFormatPlain = 1,

		 /// <summary>
		 /// SupportByVersion Outlook 10, 11, 12, 14
		 /// </summary>
		 /// <remarks>2</remarks>
		 [SupportByVersionAttribute("Outlook", 10,11,12,14)]
		 olFormatHTML = 2,

		 /// <summary>
		 /// SupportByVersion Outlook 10, 11, 12, 14
		 /// </summary>
		 /// <remarks>3</remarks>
		 [SupportByVersionAttribute("Outlook", 10,11,12,14)]
		 olFormatRichText = 3
	}
}