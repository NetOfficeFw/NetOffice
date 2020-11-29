using System;
using NetOffice;
using NetOffice.Attributes;
namespace NetOffice.WordApi.Enums
{
	 /// <summary>
	 /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
	 /// </summary>
	 ///<remarks> MSDN Online Documentation: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.WdHorizontalInVerticalType"/> </remarks>
	[SupportByVersion("Word", 9,10,11,12,14,15,16)]
	[EntityType(EntityType.IsEnum)]
	public enum WdHorizontalInVerticalType
	{
		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>0</remarks>
		 [SupportByVersion("Word", 9,10,11,12,14,15,16)]
		 wdHorizontalInVerticalNone = 0,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>1</remarks>
		 [SupportByVersion("Word", 9,10,11,12,14,15,16)]
		 wdHorizontalInVerticalFitInLine = 1,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>2</remarks>
		 [SupportByVersion("Word", 9,10,11,12,14,15,16)]
		 wdHorizontalInVerticalResizeLine = 2
	}
}