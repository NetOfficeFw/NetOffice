using System;
using NetOffice;
using NetOffice.Attributes;
namespace NetOffice.PublisherApi.Enums
{
	 /// <summary>
	 /// SupportByVersion Publisher 14, 15, 16
	 /// </summary>
	[SupportByVersion("Publisher", 14,15,16)]
	[EntityType(EntityType.IsEnum)]
	public enum PbTextOrientation
	{
		 /// <summary>
		 /// SupportByVersion Publisher 14, 15, 16
		 /// </summary>
		 /// <remarks>-2</remarks>
		 [SupportByVersion("Publisher", 14,15,16)]
		 pbTextOrientationMixed = -2,

		 /// <summary>
		 /// SupportByVersion Publisher 14, 15, 16
		 /// </summary>
		 /// <remarks>1</remarks>
		 [SupportByVersion("Publisher", 14,15,16)]
		 pbTextOrientationHorizontal = 1,

		 /// <summary>
		 /// SupportByVersion Publisher 14, 15, 16
		 /// </summary>
		 /// <remarks>2</remarks>
		 [SupportByVersion("Publisher", 14,15,16)]
		 pbTextOrientationVerticalEastAsia = 2,

		 /// <summary>
		 /// SupportByVersion Publisher 14, 15, 16
		 /// </summary>
		 /// <remarks>256</remarks>
		 [SupportByVersion("Publisher", 14,15,16)]
		 pbTextOrientationRightToLeft = 256
	}
}