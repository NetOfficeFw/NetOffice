using System;
using NetOffice;
namespace NetOffice.WordApi.Enums
{
	 /// <summary>
	 /// SupportByVersion Word 14
	 /// </summary>
	[SupportByVersionAttribute("Word", 14)]
	[EntityTypeAttribute(EntityType.IsEnum)]
	public enum XlOrientation
	{
		 /// <summary>
		 /// SupportByVersion Word 14
		 /// </summary>
		 /// <remarks>-4170</remarks>
		 [SupportByVersionAttribute("Word", 14)]
		 xlDownward = -4170,

		 /// <summary>
		 /// SupportByVersion Word 14
		 /// </summary>
		 /// <remarks>-4128</remarks>
		 [SupportByVersionAttribute("Word", 14)]
		 xlHorizontal = -4128,

		 /// <summary>
		 /// SupportByVersion Word 14
		 /// </summary>
		 /// <remarks>-4171</remarks>
		 [SupportByVersionAttribute("Word", 14)]
		 xlUpward = -4171,

		 /// <summary>
		 /// SupportByVersion Word 14
		 /// </summary>
		 /// <remarks>-4166</remarks>
		 [SupportByVersionAttribute("Word", 14)]
		 xlVertical = -4166
	}
}