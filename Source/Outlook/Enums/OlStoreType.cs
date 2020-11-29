using System;
using NetOffice;
using NetOffice.Attributes;
namespace NetOffice.OutlookApi.Enums
{
	 /// <summary>
	 /// SupportByVersion Outlook 11, 12, 14, 15, 16
	 /// </summary>
	 ///<remarks> MSDN Online Documentation: <see href="https://docs.microsoft.com/en-us/office/vba/api/Outlook.OlStoreType"/> </remarks>
	[SupportByVersion("Outlook", 11,12,14,15,16)]
	[EntityType(EntityType.IsEnum)]
	public enum OlStoreType
	{
		 /// <summary>
		 /// SupportByVersion Outlook 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>1</remarks>
		 [SupportByVersion("Outlook", 11,12,14,15,16)]
		 olStoreDefault = 1,

		 /// <summary>
		 /// SupportByVersion Outlook 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>2</remarks>
		 [SupportByVersion("Outlook", 11,12,14,15,16)]
		 olStoreUnicode = 2,

		 /// <summary>
		 /// SupportByVersion Outlook 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>3</remarks>
		 [SupportByVersion("Outlook", 11,12,14,15,16)]
		 olStoreANSI = 3
	}
}