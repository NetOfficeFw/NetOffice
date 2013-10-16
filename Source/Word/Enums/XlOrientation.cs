using System;
using NetOffice;
namespace NetOffice.WordApi.Enums
{
	 /// <summary>
	 /// SupportByVersion Word 14, 15
	 /// </summary>
	[SupportByVersionAttribute("Word", 14,15)]
	[EntityTypeAttribute(EntityType.IsEnum)]
	public enum XlOrientation
	{
		 /// <summary>
		 /// SupportByVersion Word 14, 15
		 /// </summary>
		 /// <remarks>-4170</remarks>
		 [SupportByVersionAttribute("Word", 14,15)]
		 xlDownward = -4170,

		 /// <summary>
		 /// SupportByVersion Word 14, 15
		 /// </summary>
		 /// <remarks>-4128</remarks>
		 [SupportByVersionAttribute("Word", 14,15)]
		 xlHorizontal = -4128,

		 /// <summary>
		 /// SupportByVersion Word 14, 15
		 /// </summary>
		 /// <remarks>-4171</remarks>
		 [SupportByVersionAttribute("Word", 14,15)]
		 xlUpward = -4171,

		 /// <summary>
		 /// SupportByVersion Word 14, 15
		 /// </summary>
		 /// <remarks>-4166</remarks>
		 [SupportByVersionAttribute("Word", 14,15)]
		 xlVertical = -4166
	}
}