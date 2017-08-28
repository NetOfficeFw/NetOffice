using System;
using NetOffice;
using NetOffice.Attributes;
namespace NetOffice.OfficeApi.Enums
{
	 /// <summary>
	 /// SupportByVersion Office 14, 15, 16
	 /// </summary>
	[SupportByVersion("Office", 14,15,16)]
	[EntityType(EntityType.IsEnum)]
	public enum XlPieSliceIndex
	{
		 /// <summary>
		 /// SupportByVersion Office 14, 15, 16
		 /// </summary>
		 /// <remarks>1</remarks>
		 [SupportByVersion("Office", 14,15,16)]
		 xlOuterCounterClockwisePoint = 1,

		 /// <summary>
		 /// SupportByVersion Office 14, 15, 16
		 /// </summary>
		 /// <remarks>2</remarks>
		 [SupportByVersion("Office", 14,15,16)]
		 xlOuterCenterPoint = 2,

		 /// <summary>
		 /// SupportByVersion Office 14, 15, 16
		 /// </summary>
		 /// <remarks>3</remarks>
		 [SupportByVersion("Office", 14,15,16)]
		 xlOuterClockwisePoint = 3,

		 /// <summary>
		 /// SupportByVersion Office 14, 15, 16
		 /// </summary>
		 /// <remarks>4</remarks>
		 [SupportByVersion("Office", 14,15,16)]
		 xlMidClockwiseRadiusPoint = 4,

		 /// <summary>
		 /// SupportByVersion Office 14, 15, 16
		 /// </summary>
		 /// <remarks>5</remarks>
		 [SupportByVersion("Office", 14,15,16)]
		 xlCenterPoint = 5,

		 /// <summary>
		 /// SupportByVersion Office 14, 15, 16
		 /// </summary>
		 /// <remarks>6</remarks>
		 [SupportByVersion("Office", 14,15,16)]
		 xlMidCounterClockwiseRadiusPoint = 6,

		 /// <summary>
		 /// SupportByVersion Office 14, 15, 16
		 /// </summary>
		 /// <remarks>7</remarks>
		 [SupportByVersion("Office", 14,15,16)]
		 xlInnerClockwisePoint = 7,

		 /// <summary>
		 /// SupportByVersion Office 14, 15, 16
		 /// </summary>
		 /// <remarks>8</remarks>
		 [SupportByVersion("Office", 14,15,16)]
		 xlInnerCenterPoint = 8,

		 /// <summary>
		 /// SupportByVersion Office 14, 15, 16
		 /// </summary>
		 /// <remarks>9</remarks>
		 [SupportByVersion("Office", 14,15,16)]
		 xlInnerCounterClockwisePoint = 9
	}
}