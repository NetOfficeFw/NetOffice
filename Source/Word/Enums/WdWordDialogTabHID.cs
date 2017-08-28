using System;
using NetOffice;
using NetOffice.Attributes;
namespace NetOffice.WordApi.Enums
{
	 /// <summary>
	 /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
	 /// </summary>
	[SupportByVersion("Word", 9,10,11,12,14,15,16)]
	[EntityType(EntityType.IsEnum)]
	public enum WdWordDialogTabHID
	{
		 /// <summary>
		 /// SupportByVersion Word 9
		 /// </summary>
		 /// <remarks>0</remarks>
		 [SupportByVersion("Word", 9)]
		 emptyenum = 0,

		 /// <summary>
		 /// SupportByVersion Word 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>150001</remarks>
		 [SupportByVersion("Word", 10,11,12,14,15,16)]
		 wdDialogFilePageSetupTabPaperSize = 150001,

		 /// <summary>
		 /// SupportByVersion Word 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>150002</remarks>
		 [SupportByVersion("Word", 10,11,12,14,15,16)]
		 wdDialogFilePageSetupTabPaperSource = 150002
	}
}