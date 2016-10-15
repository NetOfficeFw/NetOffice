using System;
using NetOffice;
namespace NetOffice.PowerPointApi.Enums
{
	 /// <summary>
	 /// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
	 /// </summary>
	 ///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff744288.aspx </remarks>
	[SupportByVersionAttribute("PowerPoint", 9,10,11,12,14,15,16)]
	[EntityTypeAttribute(EntityType.IsEnum)]
	public enum PpDateTimeFormat
	{
		 /// <summary>
		 /// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>-2</remarks>
		 [SupportByVersionAttribute("PowerPoint", 9,10,11,12,14,15,16)]
		 ppDateTimeFormatMixed = -2,

		 /// <summary>
		 /// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>1</remarks>
		 [SupportByVersionAttribute("PowerPoint", 9,10,11,12,14,15,16)]
		 ppDateTimeMdyy = 1,

		 /// <summary>
		 /// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>2</remarks>
		 [SupportByVersionAttribute("PowerPoint", 9,10,11,12,14,15,16)]
		 ppDateTimeddddMMMMddyyyy = 2,

		 /// <summary>
		 /// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>3</remarks>
		 [SupportByVersionAttribute("PowerPoint", 9,10,11,12,14,15,16)]
		 ppDateTimedMMMMyyyy = 3,

		 /// <summary>
		 /// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>4</remarks>
		 [SupportByVersionAttribute("PowerPoint", 9,10,11,12,14,15,16)]
		 ppDateTimeMMMMdyyyy = 4,

		 /// <summary>
		 /// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>5</remarks>
		 [SupportByVersionAttribute("PowerPoint", 9,10,11,12,14,15,16)]
		 ppDateTimedMMMyy = 5,

		 /// <summary>
		 /// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>6</remarks>
		 [SupportByVersionAttribute("PowerPoint", 9,10,11,12,14,15,16)]
		 ppDateTimeMMMMyy = 6,

		 /// <summary>
		 /// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>7</remarks>
		 [SupportByVersionAttribute("PowerPoint", 9,10,11,12,14,15,16)]
		 ppDateTimeMMyy = 7,

		 /// <summary>
		 /// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>8</remarks>
		 [SupportByVersionAttribute("PowerPoint", 9,10,11,12,14,15,16)]
		 ppDateTimeMMddyyHmm = 8,

		 /// <summary>
		 /// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>9</remarks>
		 [SupportByVersionAttribute("PowerPoint", 9,10,11,12,14,15,16)]
		 ppDateTimeMMddyyhmmAMPM = 9,

		 /// <summary>
		 /// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>10</remarks>
		 [SupportByVersionAttribute("PowerPoint", 9,10,11,12,14,15,16)]
		 ppDateTimeHmm = 10,

		 /// <summary>
		 /// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>11</remarks>
		 [SupportByVersionAttribute("PowerPoint", 9,10,11,12,14,15,16)]
		 ppDateTimeHmmss = 11,

		 /// <summary>
		 /// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>12</remarks>
		 [SupportByVersionAttribute("PowerPoint", 9,10,11,12,14,15,16)]
		 ppDateTimehmmAMPM = 12,

		 /// <summary>
		 /// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>13</remarks>
		 [SupportByVersionAttribute("PowerPoint", 9,10,11,12,14,15,16)]
		 ppDateTimehmmssAMPM = 13,

		 /// <summary>
		 /// SupportByVersion PowerPoint 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>14</remarks>
		 [SupportByVersionAttribute("PowerPoint", 10,11,12,14,15,16)]
		 ppDateTimeFigureOut = 14,

		 /// <summary>
		 /// SupportByVersion PowerPoint 15,16
		 /// </summary>
		 /// <remarks>15</remarks>
		 [SupportByVersionAttribute("PowerPoint", 15, 16)]
		 ppDateTimeUAQ1 = 15,

		 /// <summary>
		 /// SupportByVersion PowerPoint 15,16
		 /// </summary>
		 /// <remarks>16</remarks>
		 [SupportByVersionAttribute("PowerPoint", 15, 16)]
		 ppDateTimeUAQ2 = 16,

		 /// <summary>
		 /// SupportByVersion PowerPoint 15,16
		 /// </summary>
		 /// <remarks>17</remarks>
		 [SupportByVersionAttribute("PowerPoint", 15, 16)]
		 ppDateTimeUAQ3 = 17,

		 /// <summary>
		 /// SupportByVersion PowerPoint 15,16
		 /// </summary>
		 /// <remarks>18</remarks>
		 [SupportByVersionAttribute("PowerPoint", 15, 16)]
		 ppDateTimeUAQ4 = 18,

		 /// <summary>
		 /// SupportByVersion PowerPoint 15,16
		 /// </summary>
		 /// <remarks>19</remarks>
		 [SupportByVersionAttribute("PowerPoint", 15, 16)]
		 ppDateTimeUAQ5 = 19,

		 /// <summary>
		 /// SupportByVersion PowerPoint 15,16
		 /// </summary>
		 /// <remarks>20</remarks>
		 [SupportByVersionAttribute("PowerPoint", 15, 16)]
		 ppDateTimeUAQ6 = 20,

		 /// <summary>
		 /// SupportByVersion PowerPoint 15,16
		 /// </summary>
		 /// <remarks>21</remarks>
		 [SupportByVersionAttribute("PowerPoint", 15, 16)]
		 ppDateTimeUAQ7 = 21
	}
}