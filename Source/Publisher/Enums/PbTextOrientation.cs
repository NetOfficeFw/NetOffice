using System;
using NetOffice;
namespace NetOffice.PublisherApi.Enums
{
	 /// <summary>
	 /// SupportByVersion Publisher 14, 15, 16
	 /// </summary>
	[SupportByVersionAttribute("Publisher", 14,15,16)]
	[EntityTypeAttribute(EntityType.IsEnum)]
	public enum PbTextOrientation
	{
		 /// <summary>
		 /// SupportByVersion Publisher 14, 15, 16
		 /// </summary>
		 /// <remarks>-2</remarks>
		 [SupportByVersionAttribute("Publisher", 14,15,16)]
		 pbTextOrientationMixed = -2,

		 /// <summary>
		 /// SupportByVersion Publisher 14, 15, 16
		 /// </summary>
		 /// <remarks>1</remarks>
		 [SupportByVersionAttribute("Publisher", 14,15,16)]
		 pbTextOrientationHorizontal = 1,

		 /// <summary>
		 /// SupportByVersion Publisher 14, 15, 16
		 /// </summary>
		 /// <remarks>2</remarks>
		 [SupportByVersionAttribute("Publisher", 14,15,16)]
		 pbTextOrientationVerticalEastAsia = 2,

		 /// <summary>
		 /// SupportByVersion Publisher 14, 15, 16
		 /// </summary>
		 /// <remarks>256</remarks>
		 [SupportByVersionAttribute("Publisher", 14,15,16)]
		 pbTextOrientationRightToLeft = 256
	}
}