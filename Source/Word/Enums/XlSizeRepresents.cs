using System;
using NetOffice;
namespace NetOffice.WordApi.Enums
{
	 /// <summary>
	 /// SupportByVersion Word 14
	 /// </summary>
	[SupportByVersionAttribute("Word", 14)]
	[EntityTypeAttribute(EntityType.IsEnum)]
	public enum XlSizeRepresents
	{
		 /// <summary>
		 /// SupportByVersion Word 14
		 /// </summary>
		 /// <remarks>2</remarks>
		 [SupportByVersionAttribute("Word", 14)]
		 xlSizeIsWidth = 2,

		 /// <summary>
		 /// SupportByVersion Word 14
		 /// </summary>
		 /// <remarks>1</remarks>
		 [SupportByVersionAttribute("Word", 14)]
		 xlSizeIsArea = 1
	}
}