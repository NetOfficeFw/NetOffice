using System;
using NetOffice;
using NetOffice.Attributes;
namespace NetOffice.OutlookApi.Enums
{
	 /// <summary>
	 /// SupportByVersion Outlook 16
	 /// </summary>
	[SupportByVersion("Outlook", 16)]
	[EntityType(EntityType.IsEnum)]
	public enum OlUnifiedGroupType
	{
		 /// <summary>
		 /// SupportByVersion Outlook 16
		 /// </summary>
		 /// <remarks>1</remarks>
		 [SupportByVersion("Outlook", 16)]
		 PrivateGroup = 1,

		 /// <summary>
		 /// SupportByVersion Outlook 16
		 /// </summary>
		 /// <remarks>2</remarks>
		 [SupportByVersion("Outlook", 16)]
		 PublicGroup = 2
	}
}