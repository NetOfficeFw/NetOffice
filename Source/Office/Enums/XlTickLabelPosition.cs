using System;
using NetOffice;
using NetOffice.Attributes;
namespace NetOffice.OfficeApi.Enums
{
	 /// <summary>
	 /// SupportByVersion Office 12, 14, 15, 16
	 /// </summary>
	[SupportByVersion("Office", 12,14,15,16)]
	[EntityType(EntityType.IsEnum)]
	public enum XlTickLabelPosition
	{
		 /// <summary>
		 /// SupportByVersion Office 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>-4127</remarks>
		 [SupportByVersion("Office", 12,14,15,16)]
		 xlTickLabelPositionHigh = -4127,

		 /// <summary>
		 /// SupportByVersion Office 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>-4134</remarks>
		 [SupportByVersion("Office", 12,14,15,16)]
		 xlTickLabelPositionLow = -4134,

		 /// <summary>
		 /// SupportByVersion Office 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>4</remarks>
		 [SupportByVersion("Office", 12,14,15,16)]
		 xlTickLabelPositionNextToAxis = 4,

		 /// <summary>
		 /// SupportByVersion Office 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>-4142</remarks>
		 [SupportByVersion("Office", 12,14,15,16)]
		 xlTickLabelPositionNone = -4142
	}
}