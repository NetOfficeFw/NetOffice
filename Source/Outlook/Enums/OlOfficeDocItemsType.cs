using System;
using NetOffice;
namespace NetOffice.OutlookApi.Enums
{
	 /// <summary>
	 /// SupportByVersion Outlook 10, 11, 12, 14, 15
	 /// </summary>
	[SupportByVersionAttribute("Outlook", 10,11,12,14,15)]
	[EntityTypeAttribute(EntityType.IsEnum)]
	public enum OlOfficeDocItemsType
	{
		 /// <summary>
		 /// SupportByVersion Outlook 10, 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>8</remarks>
		 [SupportByVersionAttribute("Outlook", 10,11,12,14,15)]
		 olExcelWorkSheetItem = 8,

		 /// <summary>
		 /// SupportByVersion Outlook 10, 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>9</remarks>
		 [SupportByVersionAttribute("Outlook", 10,11,12,14,15)]
		 olWordDocumentItem = 9,

		 /// <summary>
		 /// SupportByVersion Outlook 10, 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>10</remarks>
		 [SupportByVersionAttribute("Outlook", 10,11,12,14,15)]
		 olPowerPointShowItem = 10
	}
}