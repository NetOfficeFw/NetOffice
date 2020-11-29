using System;
using NetOffice;
using NetOffice.Attributes;
namespace NetOffice.WordApi.Enums
{
	 /// <summary>
	 /// SupportByVersion Word 12, 14, 15, 16
	 /// </summary>
	 ///<remarks> MSDN Online Documentation: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.WdExportItem"/> </remarks>
	[SupportByVersion("Word", 12,14,15,16)]
	[EntityType(EntityType.IsEnum)]
	public enum WdExportItem
	{
		 /// <summary>
		 /// SupportByVersion Word 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>0</remarks>
		 [SupportByVersion("Word", 12,14,15,16)]
		 wdExportDocumentContent = 0,

		 /// <summary>
		 /// SupportByVersion Word 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>7</remarks>
		 [SupportByVersion("Word", 12,14,15,16)]
		 wdExportDocumentWithMarkup = 7
	}
}