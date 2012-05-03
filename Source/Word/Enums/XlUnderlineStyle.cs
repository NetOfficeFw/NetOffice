using System;
using NetOffice;
namespace NetOffice.WordApi.Enums
{
	 /// <summary>
	 /// SupportByVersion Word 14
	 /// </summary>
	[SupportByVersionAttribute("Word", 14)]
	[EntityTypeAttribute(EntityType.IsEnum)]
	public enum XlUnderlineStyle
	{
		 /// <summary>
		 /// SupportByVersion Word 14
		 /// </summary>
		 /// <remarks>-4119</remarks>
		 [SupportByVersionAttribute("Word", 14)]
		 xlUnderlineStyleDouble = -4119,

		 /// <summary>
		 /// SupportByVersion Word 14
		 /// </summary>
		 /// <remarks>5</remarks>
		 [SupportByVersionAttribute("Word", 14)]
		 xlUnderlineStyleDoubleAccounting = 5,

		 /// <summary>
		 /// SupportByVersion Word 14
		 /// </summary>
		 /// <remarks>-4142</remarks>
		 [SupportByVersionAttribute("Word", 14)]
		 xlUnderlineStyleNone = -4142,

		 /// <summary>
		 /// SupportByVersion Word 14
		 /// </summary>
		 /// <remarks>2</remarks>
		 [SupportByVersionAttribute("Word", 14)]
		 xlUnderlineStyleSingle = 2,

		 /// <summary>
		 /// SupportByVersion Word 14
		 /// </summary>
		 /// <remarks>4</remarks>
		 [SupportByVersionAttribute("Word", 14)]
		 xlUnderlineStyleSingleAccounting = 4
	}
}