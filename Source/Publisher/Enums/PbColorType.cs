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
	public enum PbColorType
	{
		 /// <summary>
		 /// SupportByVersion Publisher 14, 15, 16
		 /// </summary>
		 /// <remarks>-2</remarks>
		 [SupportByVersion("Publisher", 14,15,16)]
		 pbColorTypeMixed = -2,

		 /// <summary>
		 /// SupportByVersion Publisher 14, 15, 16
		 /// </summary>
		 /// <remarks>1</remarks>
		 [SupportByVersion("Publisher", 14,15,16)]
		 pbColorTypeRGB = 1,

		 /// <summary>
		 /// SupportByVersion Publisher 14, 15, 16
		 /// </summary>
		 /// <remarks>2</remarks>
		 [SupportByVersion("Publisher", 14,15,16)]
		 pbColorTypeScheme = 2,

		 /// <summary>
		 /// SupportByVersion Publisher 14, 15, 16
		 /// </summary>
		 /// <remarks>3</remarks>
		 [SupportByVersion("Publisher", 14,15,16)]
		 pbColorTypeCMYK = 3,

		 /// <summary>
		 /// SupportByVersion Publisher 14, 15, 16
		 /// </summary>
		 /// <remarks>4</remarks>
		 [SupportByVersion("Publisher", 14,15,16)]
		 pbColorTypeCMS = 4,

		 /// <summary>
		 /// SupportByVersion Publisher 14, 15, 16
		 /// </summary>
		 /// <remarks>5</remarks>
		 [SupportByVersion("Publisher", 14,15,16)]
		 pbColorTypeInk = 5
	}
}