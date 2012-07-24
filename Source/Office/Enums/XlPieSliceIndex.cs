using System;
using NetOffice;
namespace NetOffice.OfficeApi.Enums
{
	 /// <summary>
	 /// SupportByVersion Office 14, 15
	 /// </summary>
	[SupportByVersionAttribute("Office", 14,15)]
	[EntityTypeAttribute(EntityType.IsEnum)]
	public enum XlPieSliceIndex
	{
		 /// <summary>
		 /// SupportByVersion Office 14, 15
		 /// </summary>
		 /// <remarks>1</remarks>
		 [SupportByVersionAttribute("Office", 14,15)]
		 xlOuterCounterClockwisePoint = 1,

		 /// <summary>
		 /// SupportByVersion Office 14, 15
		 /// </summary>
		 /// <remarks>2</remarks>
		 [SupportByVersionAttribute("Office", 14,15)]
		 xlOuterCenterPoint = 2,

		 /// <summary>
		 /// SupportByVersion Office 14, 15
		 /// </summary>
		 /// <remarks>3</remarks>
		 [SupportByVersionAttribute("Office", 14,15)]
		 xlOuterClockwisePoint = 3,

		 /// <summary>
		 /// SupportByVersion Office 14, 15
		 /// </summary>
		 /// <remarks>4</remarks>
		 [SupportByVersionAttribute("Office", 14,15)]
		 xlMidClockwiseRadiusPoint = 4,

		 /// <summary>
		 /// SupportByVersion Office 14, 15
		 /// </summary>
		 /// <remarks>5</remarks>
		 [SupportByVersionAttribute("Office", 14,15)]
		 xlCenterPoint = 5,

		 /// <summary>
		 /// SupportByVersion Office 14, 15
		 /// </summary>
		 /// <remarks>6</remarks>
		 [SupportByVersionAttribute("Office", 14,15)]
		 xlMidCounterClockwiseRadiusPoint = 6,

		 /// <summary>
		 /// SupportByVersion Office 14, 15
		 /// </summary>
		 /// <remarks>7</remarks>
		 [SupportByVersionAttribute("Office", 14,15)]
		 xlInnerClockwisePoint = 7,

		 /// <summary>
		 /// SupportByVersion Office 14, 15
		 /// </summary>
		 /// <remarks>8</remarks>
		 [SupportByVersionAttribute("Office", 14,15)]
		 xlInnerCenterPoint = 8,

		 /// <summary>
		 /// SupportByVersion Office 14, 15
		 /// </summary>
		 /// <remarks>9</remarks>
		 [SupportByVersionAttribute("Office", 14,15)]
		 xlInnerCounterClockwisePoint = 9
	}
}