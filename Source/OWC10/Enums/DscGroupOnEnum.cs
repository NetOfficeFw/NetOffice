using System;
using NetOffice;
using NetOffice.Attributes;
namespace NetOffice.OWC10Api.Enums
{
	 /// <summary>
	 /// SupportByVersion OWC10 1
	 /// </summary>
	[SupportByVersion("OWC10", 1)]
	[EntityType(EntityType.IsEnum)]
	public enum DscGroupOnEnum
	{
		 /// <summary>
		 /// SupportByVersion OWC10 1
		 /// </summary>
		 /// <remarks>0</remarks>
		 [SupportByVersion("OWC10", 1)]
		 dscEachValue = 0,

		 /// <summary>
		 /// SupportByVersion OWC10 1
		 /// </summary>
		 /// <remarks>1</remarks>
		 [SupportByVersion("OWC10", 1)]
		 dscPrefix = 1,

		 /// <summary>
		 /// SupportByVersion OWC10 1
		 /// </summary>
		 /// <remarks>2</remarks>
		 [SupportByVersion("OWC10", 1)]
		 dscYear = 2,

		 /// <summary>
		 /// SupportByVersion OWC10 1
		 /// </summary>
		 /// <remarks>3</remarks>
		 [SupportByVersion("OWC10", 1)]
		 dscQuarter = 3,

		 /// <summary>
		 /// SupportByVersion OWC10 1
		 /// </summary>
		 /// <remarks>4</remarks>
		 [SupportByVersion("OWC10", 1)]
		 dscMonth = 4,

		 /// <summary>
		 /// SupportByVersion OWC10 1
		 /// </summary>
		 /// <remarks>5</remarks>
		 [SupportByVersion("OWC10", 1)]
		 dscWeek = 5,

		 /// <summary>
		 /// SupportByVersion OWC10 1
		 /// </summary>
		 /// <remarks>6</remarks>
		 [SupportByVersion("OWC10", 1)]
		 dscDay = 6,

		 /// <summary>
		 /// SupportByVersion OWC10 1
		 /// </summary>
		 /// <remarks>7</remarks>
		 [SupportByVersion("OWC10", 1)]
		 dscHour = 7,

		 /// <summary>
		 /// SupportByVersion OWC10 1
		 /// </summary>
		 /// <remarks>8</remarks>
		 [SupportByVersion("OWC10", 1)]
		 dscMinute = 8,

		 /// <summary>
		 /// SupportByVersion OWC10 1
		 /// </summary>
		 /// <remarks>9</remarks>
		 [SupportByVersion("OWC10", 1)]
		 dscInterval = 9
	}
}