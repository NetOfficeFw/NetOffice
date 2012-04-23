using System;
using NetOffice;
namespace NetOffice.OfficeApi.Enums
{
	 /// <summary>
	 /// SupportByVersion Office 12, 14
	 /// </summary>
	[SupportByVersionAttribute("Office", 12,14)]
	[EntityTypeAttribute(EntityType.IsEnum)]
	public enum XlChartOrientation
	{
		 /// <summary>
		 /// SupportByVersion Office 12, 14
		 /// </summary>
		 /// <remarks>-4170</remarks>
		 [SupportByVersionAttribute("Office", 12,14)]
		 xlDownward = -4170,

		 /// <summary>
		 /// SupportByVersion Office 12, 14
		 /// </summary>
		 /// <remarks>-4128</remarks>
		 [SupportByVersionAttribute("Office", 12,14)]
		 xlHorizontal = -4128,

		 /// <summary>
		 /// SupportByVersion Office 12, 14
		 /// </summary>
		 /// <remarks>-4171</remarks>
		 [SupportByVersionAttribute("Office", 12,14)]
		 xlUpward = -4171,

		 /// <summary>
		 /// SupportByVersion Office 12, 14
		 /// </summary>
		 /// <remarks>-4166</remarks>
		 [SupportByVersionAttribute("Office", 12,14)]
		 xlVertical = -4166
	}
}