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
	public enum XlColorIndex
	{
		 /// <summary>
		 /// SupportByVersion Office 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>-4105</remarks>
		 [SupportByVersion("Office", 12,14,15,16)]
		 xlColorIndexAutomatic = -4105,

		 /// <summary>
		 /// SupportByVersion Office 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>-4142</remarks>
		 [SupportByVersion("Office", 12,14,15,16)]
		 xlColorIndexNone = -4142
	}
}