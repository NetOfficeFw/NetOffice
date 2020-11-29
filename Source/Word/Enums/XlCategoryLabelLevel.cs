using System;
using NetOffice;
using NetOffice.Attributes;
namespace NetOffice.WordApi.Enums
{
	 /// <summary>
	 /// SupportByVersion Word 15,16
	 /// </summary>
	 ///<remarks> MSDN Online Documentation: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.xlcategorylabellevel"/> </remarks>
	[SupportByVersion("Word", 15, 16)]
	[EntityType(EntityType.IsEnum)]
	public enum XlCategoryLabelLevel
	{
		 /// <summary>
		 /// SupportByVersion Word 15,16
		 /// </summary>
		 /// <remarks>-3</remarks>
		 [SupportByVersion("Word", 15, 16)]
		 xlCategoryLabelLevelNone = -3,

		 /// <summary>
		 /// SupportByVersion Word 15,16
		 /// </summary>
		 /// <remarks>-2</remarks>
		 [SupportByVersion("Word", 15, 16)]
		 xlCategoryLabelLevelCustom = -2,

		 /// <summary>
		 /// SupportByVersion Word 15,16
		 /// </summary>
		 /// <remarks>-1</remarks>
		 [SupportByVersion("Word", 15, 16)]
		 xlCategoryLabelLevelAll = -1
	}
}