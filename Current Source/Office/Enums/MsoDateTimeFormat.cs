using System;
using NetOffice;
namespace NetOffice.OfficeApi.Enums
{
	 /// <summary>
	 /// SupportByVersion Office 12, 14
	 /// </summary>
	[SupportByVersionAttribute("Office", 12,14)]
	[EntityTypeAttribute(EntityType.IsEnum)]
	public enum MsoDateTimeFormat
	{
		 /// <summary>
		 /// SupportByVersion Office 12, 14
		 /// </summary>
		 /// <remarks>-2</remarks>
		 [SupportByVersionAttribute("Office", 12,14)]
		 msoDateTimeFormatMixed = -2,

		 /// <summary>
		 /// SupportByVersion Office 12, 14
		 /// </summary>
		 /// <remarks>1</remarks>
		 [SupportByVersionAttribute("Office", 12,14)]
		 msoDateTimeMdyy = 1,

		 /// <summary>
		 /// SupportByVersion Office 12, 14
		 /// </summary>
		 /// <remarks>2</remarks>
		 [SupportByVersionAttribute("Office", 12,14)]
		 msoDateTimeddddMMMMddyyyy = 2,

		 /// <summary>
		 /// SupportByVersion Office 12, 14
		 /// </summary>
		 /// <remarks>3</remarks>
		 [SupportByVersionAttribute("Office", 12,14)]
		 msoDateTimedMMMMyyyy = 3,

		 /// <summary>
		 /// SupportByVersion Office 12, 14
		 /// </summary>
		 /// <remarks>4</remarks>
		 [SupportByVersionAttribute("Office", 12,14)]
		 msoDateTimeMMMMdyyyy = 4,

		 /// <summary>
		 /// SupportByVersion Office 12, 14
		 /// </summary>
		 /// <remarks>5</remarks>
		 [SupportByVersionAttribute("Office", 12,14)]
		 msoDateTimedMMMyy = 5,

		 /// <summary>
		 /// SupportByVersion Office 12, 14
		 /// </summary>
		 /// <remarks>6</remarks>
		 [SupportByVersionAttribute("Office", 12,14)]
		 msoDateTimeMMMMyy = 6,

		 /// <summary>
		 /// SupportByVersion Office 12, 14
		 /// </summary>
		 /// <remarks>7</remarks>
		 [SupportByVersionAttribute("Office", 12,14)]
		 msoDateTimeMMyy = 7,

		 /// <summary>
		 /// SupportByVersion Office 12, 14
		 /// </summary>
		 /// <remarks>8</remarks>
		 [SupportByVersionAttribute("Office", 12,14)]
		 msoDateTimeMMddyyHmm = 8,

		 /// <summary>
		 /// SupportByVersion Office 12, 14
		 /// </summary>
		 /// <remarks>9</remarks>
		 [SupportByVersionAttribute("Office", 12,14)]
		 msoDateTimeMMddyyhmmAMPM = 9,

		 /// <summary>
		 /// SupportByVersion Office 12, 14
		 /// </summary>
		 /// <remarks>10</remarks>
		 [SupportByVersionAttribute("Office", 12,14)]
		 msoDateTimeHmm = 10,

		 /// <summary>
		 /// SupportByVersion Office 12, 14
		 /// </summary>
		 /// <remarks>11</remarks>
		 [SupportByVersionAttribute("Office", 12,14)]
		 msoDateTimeHmmss = 11,

		 /// <summary>
		 /// SupportByVersion Office 12, 14
		 /// </summary>
		 /// <remarks>12</remarks>
		 [SupportByVersionAttribute("Office", 12,14)]
		 msoDateTimehmmAMPM = 12,

		 /// <summary>
		 /// SupportByVersion Office 12, 14
		 /// </summary>
		 /// <remarks>13</remarks>
		 [SupportByVersionAttribute("Office", 12,14)]
		 msoDateTimehmmssAMPM = 13,

		 /// <summary>
		 /// SupportByVersion Office 12, 14
		 /// </summary>
		 /// <remarks>14</remarks>
		 [SupportByVersionAttribute("Office", 12,14)]
		 msoDateTimeFigureOut = 14
	}
}