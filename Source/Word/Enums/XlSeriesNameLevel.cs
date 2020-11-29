using System;
using NetOffice;
using NetOffice.Attributes;
namespace NetOffice.WordApi.Enums
{
	 /// <summary>
	 /// SupportByVersion Word 15,16
	 /// </summary>
	 ///<remarks> MSDN Online Documentation: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.xlseriesnamelevel"/> </remarks>
	[SupportByVersion("Word", 15, 16)]
	[EntityType(EntityType.IsEnum)]
	public enum XlSeriesNameLevel
	{
		 /// <summary>
		 /// SupportByVersion Word 15,16
		 /// </summary>
		 /// <remarks>-3</remarks>
		 [SupportByVersion("Word", 15, 16)]
		 xlSeriesNameLevelNone = -3,

		 /// <summary>
		 /// SupportByVersion Word 15,16
		 /// </summary>
		 /// <remarks>-2</remarks>
		 [SupportByVersion("Word", 15, 16)]
		 xlSeriesNameLevelCustom = -2,

		 /// <summary>
		 /// SupportByVersion Word 15,16
		 /// </summary>
		 /// <remarks>-1</remarks>
		 [SupportByVersion("Word", 15, 16)]
		 xlSeriesNameLevelAll = -1
	}
}