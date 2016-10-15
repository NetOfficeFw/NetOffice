using System;
using NetOffice;
namespace NetOffice.OfficeApi.Enums
{
	 /// <summary>
	 /// SupportByVersion Office 15,16
	 /// </summary>
	[SupportByVersionAttribute("Office", 15, 16)]
	[EntityTypeAttribute(EntityType.IsEnum)]
	public enum XlSeriesNameLevel
	{
		 /// <summary>
		 /// SupportByVersion Office 15,16
		 /// </summary>
		 /// <remarks>-3</remarks>
		 [SupportByVersionAttribute("Office", 15, 16)]
		 xlSeriesNameLevelNone = -3,

		 /// <summary>
		 /// SupportByVersion Office 15,16
		 /// </summary>
		 /// <remarks>-2</remarks>
		 [SupportByVersionAttribute("Office", 15, 16)]
		 xlSeriesNameLevelCustom = -2,

		 /// <summary>
		 /// SupportByVersion Office 15,16
		 /// </summary>
		 /// <remarks>-1</remarks>
		 [SupportByVersionAttribute("Office", 15, 16)]
		 xlSeriesNameLevelAll = -1
	}
}