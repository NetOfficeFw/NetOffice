using System;
using NetOffice;
using NetOffice.Attributes;
namespace NetOffice.WordApi.Enums
{
	 /// <summary>
	 /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
	 /// </summary>
	 ///<remarks> MSDN Online Documentation: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.WdCaptionLabelID"/> </remarks>
	[SupportByVersion("Word", 9,10,11,12,14,15,16)]
	[EntityType(EntityType.IsEnum)]
	public enum WdCaptionLabelID
	{
		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>-1</remarks>
		 [SupportByVersion("Word", 9,10,11,12,14,15,16)]
		 wdCaptionFigure = -1,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>-2</remarks>
		 [SupportByVersion("Word", 9,10,11,12,14,15,16)]
		 wdCaptionTable = -2,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>-3</remarks>
		 [SupportByVersion("Word", 9,10,11,12,14,15,16)]
		 wdCaptionEquation = -3
	}
}