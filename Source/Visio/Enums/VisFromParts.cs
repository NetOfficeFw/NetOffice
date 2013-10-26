using System;
using NetOffice;
namespace NetOffice.VisioApi.Enums
{
	 /// <summary>
	 /// SupportByVersion Visio 11, 12, 14, 15
	 /// </summary>
	[SupportByVersionAttribute("Visio", 11,12,14,15)]
	[EntityTypeAttribute(EntityType.IsEnum)]
	public enum VisFromParts
	{
		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>-1</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14,15)]
		 visConnectFromError = -1,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>0</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14,15)]
		 visFromNone = 0,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>1</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14,15)]
		 visLeftEdge = 1,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>2</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14,15)]
		 visCenterEdge = 2,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>3</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14,15)]
		 visRightEdge = 3,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>4</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14,15)]
		 visBottomEdge = 4,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>5</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14,15)]
		 visMiddleEdge = 5,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>6</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14,15)]
		 visTopEdge = 6,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>7</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14,15)]
		 visBeginX = 7,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>8</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14,15)]
		 visBeginY = 8,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>9</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14,15)]
		 visBegin = 9,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>10</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14,15)]
		 visEndX = 10,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>11</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14,15)]
		 visEndY = 11,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>12</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14,15)]
		 visEnd = 12,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>13</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14,15)]
		 visFromAngle = 13,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>14</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14,15)]
		 visFromPin = 14,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>100</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14,15)]
		 visControlPoint = 100
	}
}