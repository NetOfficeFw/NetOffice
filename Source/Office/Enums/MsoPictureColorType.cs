using System;
using NetOffice;
namespace NetOffice.OfficeApi.Enums
{
	 /// <summary>
	 /// SupportByVersion Office 9, 10, 11, 12, 14, 15
	 /// </summary>
	[SupportByVersionAttribute("Office", 9,10,11,12,14,15)]
	[EntityTypeAttribute(EntityType.IsEnum)]
	public enum MsoPictureColorType
	{
		 /// <summary>
		 /// SupportByVersion Office 9, 10, 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>-2</remarks>
		 [SupportByVersionAttribute("Office", 9,10,11,12,14,15)]
		 msoPictureMixed = -2,

		 /// <summary>
		 /// SupportByVersion Office 9, 10, 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>1</remarks>
		 [SupportByVersionAttribute("Office", 9,10,11,12,14,15)]
		 msoPictureAutomatic = 1,

		 /// <summary>
		 /// SupportByVersion Office 9, 10, 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>2</remarks>
		 [SupportByVersionAttribute("Office", 9,10,11,12,14,15)]
		 msoPictureGrayscale = 2,

		 /// <summary>
		 /// SupportByVersion Office 9, 10, 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>3</remarks>
		 [SupportByVersionAttribute("Office", 9,10,11,12,14,15)]
		 msoPictureBlackAndWhite = 3,

		 /// <summary>
		 /// SupportByVersion Office 9, 10, 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>4</remarks>
		 [SupportByVersionAttribute("Office", 9,10,11,12,14,15)]
		 msoPictureWatermark = 4
	}
}