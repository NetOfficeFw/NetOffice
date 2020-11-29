﻿using System;
using NetOffice;
using NetOffice.Attributes;
namespace NetOffice.PowerPointApi.Enums
{
	 /// <summary>
	 /// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
	 /// </summary>
	 ///<remarks> MSDN Online Documentation: <see href="https://docs.microsoft.com/en-us/office/vba/api/PowerPoint.PpDateTimeFormat"/> </remarks>
	[SupportByVersion("PowerPoint", 9,10,11,12,14,15,16)]
	[EntityType(EntityType.IsEnum)]
	public enum PpDateTimeFormat
	{
		 /// <summary>
		 /// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>-2</remarks>
		 [SupportByVersion("PowerPoint", 9,10,11,12,14,15,16)]
		 ppDateTimeFormatMixed = -2,

		 /// <summary>
		 /// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>1</remarks>
		 [SupportByVersion("PowerPoint", 9,10,11,12,14,15,16)]
		 ppDateTimeMdyy = 1,

		 /// <summary>
		 /// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>2</remarks>
		 [SupportByVersion("PowerPoint", 9,10,11,12,14,15,16)]
		 ppDateTimeddddMMMMddyyyy = 2,

		 /// <summary>
		 /// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>3</remarks>
		 [SupportByVersion("PowerPoint", 9,10,11,12,14,15,16)]
		 ppDateTimedMMMMyyyy = 3,

		 /// <summary>
		 /// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>4</remarks>
		 [SupportByVersion("PowerPoint", 9,10,11,12,14,15,16)]
		 ppDateTimeMMMMdyyyy = 4,

		 /// <summary>
		 /// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>5</remarks>
		 [SupportByVersion("PowerPoint", 9,10,11,12,14,15,16)]
		 ppDateTimedMMMyy = 5,

		 /// <summary>
		 /// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>6</remarks>
		 [SupportByVersion("PowerPoint", 9,10,11,12,14,15,16)]
		 ppDateTimeMMMMyy = 6,

		 /// <summary>
		 /// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>7</remarks>
		 [SupportByVersion("PowerPoint", 9,10,11,12,14,15,16)]
		 ppDateTimeMMyy = 7,

		 /// <summary>
		 /// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>8</remarks>
		 [SupportByVersion("PowerPoint", 9,10,11,12,14,15,16)]
		 ppDateTimeMMddyyHmm = 8,

		 /// <summary>
		 /// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>9</remarks>
		 [SupportByVersion("PowerPoint", 9,10,11,12,14,15,16)]
		 ppDateTimeMMddyyhmmAMPM = 9,

		 /// <summary>
		 /// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>10</remarks>
		 [SupportByVersion("PowerPoint", 9,10,11,12,14,15,16)]
		 ppDateTimeHmm = 10,

		 /// <summary>
		 /// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>11</remarks>
		 [SupportByVersion("PowerPoint", 9,10,11,12,14,15,16)]
		 ppDateTimeHmmss = 11,

		 /// <summary>
		 /// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>12</remarks>
		 [SupportByVersion("PowerPoint", 9,10,11,12,14,15,16)]
		 ppDateTimehmmAMPM = 12,

		 /// <summary>
		 /// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>13</remarks>
		 [SupportByVersion("PowerPoint", 9,10,11,12,14,15,16)]
		 ppDateTimehmmssAMPM = 13,

		 /// <summary>
		 /// SupportByVersion PowerPoint 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>14</remarks>
		 [SupportByVersion("PowerPoint", 10,11,12,14,15,16)]
		 ppDateTimeFigureOut = 14,

		 /// <summary>
		 /// SupportByVersion PowerPoint 15,16
		 /// </summary>
		 /// <remarks>15</remarks>
		 [SupportByVersion("PowerPoint", 15, 16)]
		 ppDateTimeUAQ1 = 15,

		 /// <summary>
		 /// SupportByVersion PowerPoint 15,16
		 /// </summary>
		 /// <remarks>16</remarks>
		 [SupportByVersion("PowerPoint", 15, 16)]
		 ppDateTimeUAQ2 = 16,

		 /// <summary>
		 /// SupportByVersion PowerPoint 15,16
		 /// </summary>
		 /// <remarks>17</remarks>
		 [SupportByVersion("PowerPoint", 15, 16)]
		 ppDateTimeUAQ3 = 17,

		 /// <summary>
		 /// SupportByVersion PowerPoint 15,16
		 /// </summary>
		 /// <remarks>18</remarks>
		 [SupportByVersion("PowerPoint", 15, 16)]
		 ppDateTimeUAQ4 = 18,

		 /// <summary>
		 /// SupportByVersion PowerPoint 15,16
		 /// </summary>
		 /// <remarks>19</remarks>
		 [SupportByVersion("PowerPoint", 15, 16)]
		 ppDateTimeUAQ5 = 19,

		 /// <summary>
		 /// SupportByVersion PowerPoint 15,16
		 /// </summary>
		 /// <remarks>20</remarks>
		 [SupportByVersion("PowerPoint", 15, 16)]
		 ppDateTimeUAQ6 = 20,

		 /// <summary>
		 /// SupportByVersion PowerPoint 15,16
		 /// </summary>
		 /// <remarks>21</remarks>
		 [SupportByVersion("PowerPoint", 15, 16)]
		 ppDateTimeUAQ7 = 21
	}
}