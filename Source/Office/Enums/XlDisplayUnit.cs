using System;
using NetOffice;
using NetOffice.Attributes;
namespace NetOffice.OfficeApi.Enums
{
	 /// <summary>
	 /// SupportByVersion Office 12, 14, 15, 16
	 /// </summary>
	 ///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff860551.aspx </remarks>
	[SupportByVersion("Office", 12,14,15,16)]
	[EntityType(EntityType.IsEnum)]
	public enum XlDisplayUnit
	{
		 /// <summary>
		 /// SupportByVersion Office 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>-2</remarks>
		 [SupportByVersion("Office", 12,14,15,16)]
		 xlHundreds = -2,

		 /// <summary>
		 /// SupportByVersion Office 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>-3</remarks>
		 [SupportByVersion("Office", 12,14,15,16)]
		 xlThousands = -3,

		 /// <summary>
		 /// SupportByVersion Office 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>-4</remarks>
		 [SupportByVersion("Office", 12,14,15,16)]
		 xlTenThousands = -4,

		 /// <summary>
		 /// SupportByVersion Office 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>-5</remarks>
		 [SupportByVersion("Office", 12,14,15,16)]
		 xlHundredThousands = -5,

		 /// <summary>
		 /// SupportByVersion Office 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>-6</remarks>
		 [SupportByVersion("Office", 12,14,15,16)]
		 xlMillions = -6,

		 /// <summary>
		 /// SupportByVersion Office 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>-7</remarks>
		 [SupportByVersion("Office", 12,14,15,16)]
		 xlTenMillions = -7,

		 /// <summary>
		 /// SupportByVersion Office 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>-8</remarks>
		 [SupportByVersion("Office", 12,14,15,16)]
		 xlHundredMillions = -8,

		 /// <summary>
		 /// SupportByVersion Office 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>-9</remarks>
		 [SupportByVersion("Office", 12,14,15,16)]
		 xlThousandMillions = -9,

		 /// <summary>
		 /// SupportByVersion Office 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>-10</remarks>
		 [SupportByVersion("Office", 12,14,15,16)]
		 xlMillionMillions = -10,

		 /// <summary>
		 /// SupportByVersion Office 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>-4114</remarks>
		 [SupportByVersion("Office", 12,14,15,16)]
		 xlDisplayUnitCustom = -4114,

		 /// <summary>
		 /// SupportByVersion Office 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>-4142</remarks>
		 [SupportByVersion("Office", 12,14,15,16)]
		 xlDisplayUnitNone = -4142
	}
}