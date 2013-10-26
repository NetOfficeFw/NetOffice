using System;
using NetOffice;
namespace NetOffice.WordApi.Enums
{
	 /// <summary>
	 /// SupportByVersion Word 15
	 /// </summary>
	 ///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/jj232146.aspx </remarks>
	[SupportByVersionAttribute("Word", 15)]
	[EntityTypeAttribute(EntityType.IsEnum)]
	public enum XlSeriesNameLevel
	{
		 /// <summary>
		 /// SupportByVersion Word 15
		 /// </summary>
		 /// <remarks>-3</remarks>
		 [SupportByVersionAttribute("Word", 15)]
		 xlSeriesNameLevelNone = -3,

		 /// <summary>
		 /// SupportByVersion Word 15
		 /// </summary>
		 /// <remarks>-2</remarks>
		 [SupportByVersionAttribute("Word", 15)]
		 xlSeriesNameLevelCustom = -2,

		 /// <summary>
		 /// SupportByVersion Word 15
		 /// </summary>
		 /// <remarks>-1</remarks>
		 [SupportByVersionAttribute("Word", 15)]
		 xlSeriesNameLevelAll = -1
	}
}