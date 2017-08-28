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
	public enum PivotFieldTypeEnum
	{
		 /// <summary>
		 /// SupportByVersion OWC10 1
		 /// </summary>
		 /// <remarks>1</remarks>
		 [SupportByVersion("OWC10", 1)]
		 plTypeRegular = 1,

		 /// <summary>
		 /// SupportByVersion OWC10 1
		 /// </summary>
		 /// <remarks>2</remarks>
		 [SupportByVersion("OWC10", 1)]
		 plTypeCalculated = 2,

		 /// <summary>
		 /// SupportByVersion OWC10 1
		 /// </summary>
		 /// <remarks>4</remarks>
		 [SupportByVersion("OWC10", 1)]
		 plTypeTimeYears = 4,

		 /// <summary>
		 /// SupportByVersion OWC10 1
		 /// </summary>
		 /// <remarks>5</remarks>
		 [SupportByVersion("OWC10", 1)]
		 plTypeTimeHalfYears = 5,

		 /// <summary>
		 /// SupportByVersion OWC10 1
		 /// </summary>
		 /// <remarks>6</remarks>
		 [SupportByVersion("OWC10", 1)]
		 plTypeTimeQuarters = 6,

		 /// <summary>
		 /// SupportByVersion OWC10 1
		 /// </summary>
		 /// <remarks>7</remarks>
		 [SupportByVersion("OWC10", 1)]
		 plTypeTimeMonths = 7,

		 /// <summary>
		 /// SupportByVersion OWC10 1
		 /// </summary>
		 /// <remarks>8</remarks>
		 [SupportByVersion("OWC10", 1)]
		 plTypeTimeWeeks = 8,

		 /// <summary>
		 /// SupportByVersion OWC10 1
		 /// </summary>
		 /// <remarks>9</remarks>
		 [SupportByVersion("OWC10", 1)]
		 plTypeTimeDays = 9,

		 /// <summary>
		 /// SupportByVersion OWC10 1
		 /// </summary>
		 /// <remarks>10</remarks>
		 [SupportByVersion("OWC10", 1)]
		 plTypeTimeHours = 10,

		 /// <summary>
		 /// SupportByVersion OWC10 1
		 /// </summary>
		 /// <remarks>11</remarks>
		 [SupportByVersion("OWC10", 1)]
		 plTypeTimeMinutes = 11,

		 /// <summary>
		 /// SupportByVersion OWC10 1
		 /// </summary>
		 /// <remarks>12</remarks>
		 [SupportByVersion("OWC10", 1)]
		 plTypeTimeSeconds = 12,

		 /// <summary>
		 /// SupportByVersion OWC10 1
		 /// </summary>
		 /// <remarks>13</remarks>
		 [SupportByVersion("OWC10", 1)]
		 plTypeTimeUndefined = 13,

		 /// <summary>
		 /// SupportByVersion OWC10 1
		 /// </summary>
		 /// <remarks>14</remarks>
		 [SupportByVersion("OWC10", 1)]
		 plTypeUnknown = 14,

		 /// <summary>
		 /// SupportByVersion OWC10 1
		 /// </summary>
		 /// <remarks>15</remarks>
		 [SupportByVersion("OWC10", 1)]
		 plTypeUserDefined = 15,

		 /// <summary>
		 /// SupportByVersion OWC10 1
		 /// </summary>
		 /// <remarks>16</remarks>
		 [SupportByVersion("OWC10", 1)]
		 plTypeTimeWeekdays = 16,

		 /// <summary>
		 /// SupportByVersion OWC10 1
		 /// </summary>
		 /// <remarks>17</remarks>
		 [SupportByVersion("OWC10", 1)]
		 plTypeCustomGroup = 17
	}
}