using System;
using NetOffice;
namespace NetOffice.PowerPointApi.Enums
{
	 /// <summary>
	 /// SupportByVersion PowerPoint 14, 15, 16
	 /// </summary>
	 ///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff744224.aspx </remarks>
	[SupportByVersionAttribute("PowerPoint", 14,15,16)]
	[EntityTypeAttribute(EntityType.IsEnum)]
	public enum XlDisplayUnit
	{
		 /// <summary>
		 /// SupportByVersion PowerPoint 14, 15, 16
		 /// </summary>
		 /// <remarks>-2</remarks>
		 [SupportByVersionAttribute("PowerPoint", 14,15,16)]
		 xlHundreds = -2,

		 /// <summary>
		 /// SupportByVersion PowerPoint 14, 15, 16
		 /// </summary>
		 /// <remarks>-3</remarks>
		 [SupportByVersionAttribute("PowerPoint", 14,15,16)]
		 xlThousands = -3,

		 /// <summary>
		 /// SupportByVersion PowerPoint 14, 15, 16
		 /// </summary>
		 /// <remarks>-4</remarks>
		 [SupportByVersionAttribute("PowerPoint", 14,15,16)]
		 xlTenThousands = -4,

		 /// <summary>
		 /// SupportByVersion PowerPoint 14, 15, 16
		 /// </summary>
		 /// <remarks>-5</remarks>
		 [SupportByVersionAttribute("PowerPoint", 14,15,16)]
		 xlHundredThousands = -5,

		 /// <summary>
		 /// SupportByVersion PowerPoint 14, 15, 16
		 /// </summary>
		 /// <remarks>-6</remarks>
		 [SupportByVersionAttribute("PowerPoint", 14,15,16)]
		 xlMillions = -6,

		 /// <summary>
		 /// SupportByVersion PowerPoint 14, 15, 16
		 /// </summary>
		 /// <remarks>-7</remarks>
		 [SupportByVersionAttribute("PowerPoint", 14,15,16)]
		 xlTenMillions = -7,

		 /// <summary>
		 /// SupportByVersion PowerPoint 14, 15, 16
		 /// </summary>
		 /// <remarks>-8</remarks>
		 [SupportByVersionAttribute("PowerPoint", 14,15,16)]
		 xlHundredMillions = -8,

		 /// <summary>
		 /// SupportByVersion PowerPoint 14, 15, 16
		 /// </summary>
		 /// <remarks>-9</remarks>
		 [SupportByVersionAttribute("PowerPoint", 14,15,16)]
		 xlThousandMillions = -9,

		 /// <summary>
		 /// SupportByVersion PowerPoint 14, 15, 16
		 /// </summary>
		 /// <remarks>-10</remarks>
		 [SupportByVersionAttribute("PowerPoint", 14,15,16)]
		 xlMillionMillions = -10
	}
}