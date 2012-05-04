using System;
using NetOffice;
namespace NetOffice.VisioApi.Enums
{
	 /// <summary>
	 /// SupportByVersion Visio 11, 12, 14
	 /// </summary>
	[SupportByVersionAttribute("Visio", 11,12,14)]
	[EntityTypeAttribute(EntityType.IsEnum)]
	public enum VisDrawSplineFlags
	{
		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14
		 /// </summary>
		 /// <remarks>1</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14)]
		 visSplinePeriodic = 1,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14
		 /// </summary>
		 /// <remarks>2</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14)]
		 visSplineDoCircles = 2,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14
		 /// </summary>
		 /// <remarks>4</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14)]
		 visSplineAbrupt = 4,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14
		 /// </summary>
		 /// <remarks>8</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14)]
		 visSpline1D = 8,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14
		 /// </summary>
		 /// <remarks>8</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14)]
		 visPolyline1D = 8,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14
		 /// </summary>
		 /// <remarks>256</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14)]
		 visPolyarcs = 256
	}
}