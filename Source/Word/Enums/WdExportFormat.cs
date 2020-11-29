using System;
using NetOffice;
using NetOffice.Attributes;
namespace NetOffice.WordApi.Enums
{
	 /// <summary>
	 /// SupportByVersion Word 12, 14, 15, 16
	 /// </summary>
	 ///<remarks> MSDN Online Documentation: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.WdExportFormat"/> </remarks>
	[SupportByVersion("Word", 12,14,15,16)]
	[EntityType(EntityType.IsEnum)]
	public enum WdExportFormat
	{
		 /// <summary>
		 /// SupportByVersion Word 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>17</remarks>
		 [SupportByVersion("Word", 12,14,15,16)]
		 wdExportFormatPDF = 17,

		 /// <summary>
		 /// SupportByVersion Word 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>18</remarks>
		 [SupportByVersion("Word", 12,14,15,16)]
		 wdExportFormatXPS = 18
	}
}