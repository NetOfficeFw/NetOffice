using System;
using NetOffice;
namespace NetOffice.ExcelApi.Enums
{
	 /// <summary>
	 /// SupportByVersion Excel 14
	 /// </summary>
	[SupportByVersionAttribute("Excel", 14)]
	[EntityTypeAttribute(EntityType.IsEnum)]
	public enum XlSpanishModes
	{
		 /// <summary>
		 /// SupportByVersion Excel 14
		 /// </summary>
		 /// <remarks>0</remarks>
		 [SupportByVersionAttribute("Excel", 14)]
		 xlSpanishTuteoOnly = 0,

		 /// <summary>
		 /// SupportByVersion Excel 14
		 /// </summary>
		 /// <remarks>1</remarks>
		 [SupportByVersionAttribute("Excel", 14)]
		 xlSpanishTuteoAndVoseo = 1,

		 /// <summary>
		 /// SupportByVersion Excel 14
		 /// </summary>
		 /// <remarks>2</remarks>
		 [SupportByVersionAttribute("Excel", 14)]
		 xlSpanishVoseoOnly = 2
	}
}