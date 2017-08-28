using System;
using NetOffice;
using NetOffice.Attributes;
namespace NetOffice.VisioApi.Enums
{
	 /// <summary>
	 /// SupportByVersion Visio 12, 14, 15, 16
	 /// </summary>
	[SupportByVersion("Visio", 12,14,15,16)]
	[EntityType(EntityType.IsEnum)]
	public enum VisGraphicPositionVertical
	{
		 /// <summary>
		 /// SupportByVersion Visio 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>0</remarks>
		 [SupportByVersion("Visio", 12,14,15,16)]
		 visGraphicBelow = 0,

		 /// <summary>
		 /// SupportByVersion Visio 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>1</remarks>
		 [SupportByVersion("Visio", 12,14,15,16)]
		 visGraphicBottomEdge = 1,

		 /// <summary>
		 /// SupportByVersion Visio 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>2</remarks>
		 [SupportByVersion("Visio", 12,14,15,16)]
		 visGraphicBottom = 2,

		 /// <summary>
		 /// SupportByVersion Visio 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>3</remarks>
		 [SupportByVersion("Visio", 12,14,15,16)]
		 visGraphicMiddle = 3,

		 /// <summary>
		 /// SupportByVersion Visio 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>4</remarks>
		 [SupportByVersion("Visio", 12,14,15,16)]
		 visGraphicTop = 4,

		 /// <summary>
		 /// SupportByVersion Visio 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>5</remarks>
		 [SupportByVersion("Visio", 12,14,15,16)]
		 visGraphicTopEdge = 5,

		 /// <summary>
		 /// SupportByVersion Visio 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>6</remarks>
		 [SupportByVersion("Visio", 12,14,15,16)]
		 visGraphicAbove = 6
	}
}