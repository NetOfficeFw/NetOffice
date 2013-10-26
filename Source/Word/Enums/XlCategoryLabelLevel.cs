using System;
using NetOffice;
namespace NetOffice.WordApi.Enums
{
	 /// <summary>
	 /// SupportByVersion Word 15
	 /// </summary>
	 ///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/jj230976.aspx </remarks>
	[SupportByVersionAttribute("Word", 15)]
	[EntityTypeAttribute(EntityType.IsEnum)]
	public enum XlCategoryLabelLevel
	{
		 /// <summary>
		 /// SupportByVersion Word 15
		 /// </summary>
		 /// <remarks>-3</remarks>
		 [SupportByVersionAttribute("Word", 15)]
		 xlCategoryLabelLevelNone = -3,

		 /// <summary>
		 /// SupportByVersion Word 15
		 /// </summary>
		 /// <remarks>-2</remarks>
		 [SupportByVersionAttribute("Word", 15)]
		 xlCategoryLabelLevelCustom = -2,

		 /// <summary>
		 /// SupportByVersion Word 15
		 /// </summary>
		 /// <remarks>-1</remarks>
		 [SupportByVersionAttribute("Word", 15)]
		 xlCategoryLabelLevelAll = -1
	}
}