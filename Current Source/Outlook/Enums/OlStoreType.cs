using System;
using LateBindingApi.Core;
namespace NetOffice.OutlookApi.Enums
{
	 /// <summary>
	 /// SupportByVersion Outlook 11, 12, 14
	 /// </summary>
	[SupportByVersionAttribute("Outlook", 11,12,14)]
	[EntityTypeAttribute(EntityType.IsEnum)]
	public enum OlStoreType
	{
		 /// <summary>
		 /// SupportByVersion Outlook 11, 12, 14
		 /// </summary>
		 /// <remarks>1</remarks>
		 [SupportByVersionAttribute("Outlook", 11,12,14)]
		 olStoreDefault = 1,

		 /// <summary>
		 /// SupportByVersion Outlook 11, 12, 14
		 /// </summary>
		 /// <remarks>2</remarks>
		 [SupportByVersionAttribute("Outlook", 11,12,14)]
		 olStoreUnicode = 2,

		 /// <summary>
		 /// SupportByVersion Outlook 11, 12, 14
		 /// </summary>
		 /// <remarks>3</remarks>
		 [SupportByVersionAttribute("Outlook", 11,12,14)]
		 olStoreANSI = 3
	}
}