using System;
using NetOffice;
using NetOffice.Attributes;
namespace NetOffice.WordApi.Enums
{
	 /// <summary>
	 /// SupportByVersion Word 12, 14, 15, 16
	 /// </summary>
	 ///<remarks> MSDN Online Documentation: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.WdPasteOptions"/> </remarks>
	[SupportByVersion("Word", 12,14,15,16)]
	[EntityType(EntityType.IsEnum)]
	public enum WdPasteOptions
	{
		 /// <summary>
		 /// SupportByVersion Word 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>0</remarks>
		 [SupportByVersion("Word", 12,14,15,16)]
		 wdKeepSourceFormatting = 0,

		 /// <summary>
		 /// SupportByVersion Word 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>1</remarks>
		 [SupportByVersion("Word", 12,14,15,16)]
		 wdMatchDestinationFormatting = 1,

		 /// <summary>
		 /// SupportByVersion Word 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>2</remarks>
		 [SupportByVersion("Word", 12,14,15,16)]
		 wdKeepTextOnly = 2,

		 /// <summary>
		 /// SupportByVersion Word 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>3</remarks>
		 [SupportByVersion("Word", 12,14,15,16)]
		 wdUseDestinationStyles = 3
	}
}