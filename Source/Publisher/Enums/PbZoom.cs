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
	public enum PbZoom
	{
		 /// <summary>
		 /// SupportByVersion Publisher 14, 15, 16
		 /// </summary>
		 /// <remarks>-1</remarks>
		 [SupportByVersion("Publisher", 14,15,16)]
		 pbZoomPageWidth = -1,

		 /// <summary>
		 /// SupportByVersion Publisher 14, 15, 16
		 /// </summary>
		 /// <remarks>-2</remarks>
		 [SupportByVersion("Publisher", 14,15,16)]
		 pbZoomWholePage = -2,

		 /// <summary>
		 /// SupportByVersion Publisher 14, 15, 16
		 /// </summary>
		 /// <remarks>-3</remarks>
		 [SupportByVersion("Publisher", 14,15,16)]
		 pbZoomFitSelection = -3
	}
}