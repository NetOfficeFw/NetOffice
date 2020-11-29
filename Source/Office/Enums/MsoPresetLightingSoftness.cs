using System;
using NetOffice;
using NetOffice.Attributes;
namespace NetOffice.OfficeApi.Enums
{
	 /// <summary>
	 /// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
	 /// </summary>
	 ///<remarks> MSDN Online Documentation: <see href="https://docs.microsoft.com/en-us/office/vba/api/Office.MsoPresetLightingSoftness"/> </remarks>
	[SupportByVersion("Office", 9,10,11,12,14,15,16)]
	[EntityType(EntityType.IsEnum)]
	public enum MsoPresetLightingSoftness
	{
		 /// <summary>
		 /// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>-2</remarks>
		 [SupportByVersion("Office", 9,10,11,12,14,15,16)]
		 msoPresetLightingSoftnessMixed = -2,

		 /// <summary>
		 /// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>1</remarks>
		 [SupportByVersion("Office", 9,10,11,12,14,15,16)]
		 msoLightingDim = 1,

		 /// <summary>
		 /// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>2</remarks>
		 [SupportByVersion("Office", 9,10,11,12,14,15,16)]
		 msoLightingNormal = 2,

		 /// <summary>
		 /// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>3</remarks>
		 [SupportByVersion("Office", 9,10,11,12,14,15,16)]
		 msoLightingBright = 3
	}
}