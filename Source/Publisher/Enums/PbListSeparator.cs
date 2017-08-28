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
	public enum PbListSeparator
	{
		 /// <summary>
		 /// SupportByVersion Publisher 14, 15, 16
		 /// </summary>
		 /// <remarks>0</remarks>
		 [SupportByVersion("Publisher", 14,15,16)]
		 pbListSeparatorParenthesis = 0,

		 /// <summary>
		 /// SupportByVersion Publisher 14, 15, 16
		 /// </summary>
		 /// <remarks>65536</remarks>
		 [SupportByVersion("Publisher", 14,15,16)]
		 pbListSeparatorDoubleParen = 65536,

		 /// <summary>
		 /// SupportByVersion Publisher 14, 15, 16
		 /// </summary>
		 /// <remarks>131072</remarks>
		 [SupportByVersion("Publisher", 14,15,16)]
		 pbListSeparatorPeriod = 131072,

		 /// <summary>
		 /// SupportByVersion Publisher 14, 15, 16
		 /// </summary>
		 /// <remarks>196608</remarks>
		 [SupportByVersion("Publisher", 14,15,16)]
		 pbListSeparatorPlain = 196608,

		 /// <summary>
		 /// SupportByVersion Publisher 14, 15, 16
		 /// </summary>
		 /// <remarks>262144</remarks>
		 [SupportByVersion("Publisher", 14,15,16)]
		 pbListSeparatorSquare = 262144,

		 /// <summary>
		 /// SupportByVersion Publisher 14, 15, 16
		 /// </summary>
		 /// <remarks>327680</remarks>
		 [SupportByVersion("Publisher", 14,15,16)]
		 pbListSeparatorColon = 327680,

		 /// <summary>
		 /// SupportByVersion Publisher 14, 15, 16
		 /// </summary>
		 /// <remarks>393216</remarks>
		 [SupportByVersion("Publisher", 14,15,16)]
		 pbListSeparatorDoubleSquare = 393216,

		 /// <summary>
		 /// SupportByVersion Publisher 14, 15, 16
		 /// </summary>
		 /// <remarks>458752</remarks>
		 [SupportByVersion("Publisher", 14,15,16)]
		 pbListSeparatorDoubleHyphen = 458752,

		 /// <summary>
		 /// SupportByVersion Publisher 14, 15, 16
		 /// </summary>
		 /// <remarks>524288</remarks>
		 [SupportByVersion("Publisher", 14,15,16)]
		 pbListSeparatorWideComma = 524288
	}
}