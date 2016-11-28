using System;
using NetOffice;
namespace NetOffice.OWC10Api.Enums
{
	 /// <summary>
	 /// SupportByVersion OWC10 1
	 /// </summary>
	[SupportByVersionAttribute("OWC10", 1)]
	[EntityTypeAttribute(EntityType.IsEnum)]
	public enum PivotFieldTypeEnum
	{
		 /// <summary>
		 /// SupportByVersion OWC10 1
		 /// </summary>
		 /// <remarks>1</remarks>
		 [SupportByVersionAttribute("OWC10", 1)]
		 plTypeRegular = 1,

		 /// <summary>
		 /// SupportByVersion OWC10 1
		 /// </summary>
		 /// <remarks>2</remarks>
		 [SupportByVersionAttribute("OWC10", 1)]
		 plTypeCalculated = 2,

		 /// <summary>
		 /// SupportByVersion OWC10 1
		 /// </summary>
		 /// <remarks>4</remarks>
		 [SupportByVersionAttribute("OWC10", 1)]
		 plTypeTimeYears = 4,

		 /// <summary>
		 /// SupportByVersion OWC10 1
		 /// </summary>
		 /// <remarks>5</remarks>
		 [SupportByVersionAttribute("OWC10", 1)]
		 plTypeTimeHalfYears = 5,

		 /// <summary>
		 /// SupportByVersion OWC10 1
		 /// </summary>
		 /// <remarks>6</remarks>
		 [SupportByVersionAttribute("OWC10", 1)]
		 plTypeTimeQuarters = 6,

		 /// <summary>
		 /// SupportByVersion OWC10 1
		 /// </summary>
		 /// <remarks>7</remarks>
		 [SupportByVersionAttribute("OWC10", 1)]
		 plTypeTimeMonths = 7,

		 /// <summary>
		 /// SupportByVersion OWC10 1
		 /// </summary>
		 /// <remarks>8</remarks>
		 [SupportByVersionAttribute("OWC10", 1)]
		 plTypeTimeWeeks = 8,

		 /// <summary>
		 /// SupportByVersion OWC10 1
		 /// </summary>
		 /// <remarks>9</remarks>
		 [SupportByVersionAttribute("OWC10", 1)]
		 plTypeTimeDays = 9,

		 /// <summary>
		 /// SupportByVersion OWC10 1
		 /// </summary>
		 /// <remarks>10</remarks>
		 [SupportByVersionAttribute("OWC10", 1)]
		 plTypeTimeHours = 10,

		 /// <summary>
		 /// SupportByVersion OWC10 1
		 /// </summary>
		 /// <remarks>11</remarks>
		 [SupportByVersionAttribute("OWC10", 1)]
		 plTypeTimeMinutes = 11,

		 /// <summary>
		 /// SupportByVersion OWC10 1
		 /// </summary>
		 /// <remarks>12</remarks>
		 [SupportByVersionAttribute("OWC10", 1)]
		 plTypeTimeSeconds = 12,

		 /// <summary>
		 /// SupportByVersion OWC10 1
		 /// </summary>
		 /// <remarks>13</remarks>
		 [SupportByVersionAttribute("OWC10", 1)]
		 plTypeTimeUndefined = 13,

		 /// <summary>
		 /// SupportByVersion OWC10 1
		 /// </summary>
		 /// <remarks>14</remarks>
		 [SupportByVersionAttribute("OWC10", 1)]
		 plTypeUnknown = 14,

		 /// <summary>
		 /// SupportByVersion OWC10 1
		 /// </summary>
		 /// <remarks>15</remarks>
		 [SupportByVersionAttribute("OWC10", 1)]
		 plTypeUserDefined = 15,

		 /// <summary>
		 /// SupportByVersion OWC10 1
		 /// </summary>
		 /// <remarks>16</remarks>
		 [SupportByVersionAttribute("OWC10", 1)]
		 plTypeTimeWeekdays = 16,

		 /// <summary>
		 /// SupportByVersion OWC10 1
		 /// </summary>
		 /// <remarks>17</remarks>
		 [SupportByVersionAttribute("OWC10", 1)]
		 plTypeCustomGroup = 17
	}
}