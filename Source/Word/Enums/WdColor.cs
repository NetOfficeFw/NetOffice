using System;
using NetOffice;
namespace NetOffice.WordApi.Enums
{
	 /// <summary>
	 /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
	 /// </summary>
	 ///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff196272.aspx </remarks>
	[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
	[EntityTypeAttribute(EntityType.IsEnum)]
	public enum WdColor
	{
		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>-16777216</remarks>
		 [SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		 wdColorAutomatic = -16777216,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>0</remarks>
		 [SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		 wdColorBlack = 0,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>16711680</remarks>
		 [SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		 wdColorBlue = 16711680,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>16776960</remarks>
		 [SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		 wdColorTurquoise = 16776960,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>65280</remarks>
		 [SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		 wdColorBrightGreen = 65280,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>16711935</remarks>
		 [SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		 wdColorPink = 16711935,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>255</remarks>
		 [SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		 wdColorRed = 255,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>65535</remarks>
		 [SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		 wdColorYellow = 65535,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>16777215</remarks>
		 [SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		 wdColorWhite = 16777215,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>8388608</remarks>
		 [SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		 wdColorDarkBlue = 8388608,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>8421376</remarks>
		 [SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		 wdColorTeal = 8421376,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>32768</remarks>
		 [SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		 wdColorGreen = 32768,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>8388736</remarks>
		 [SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		 wdColorViolet = 8388736,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>128</remarks>
		 [SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		 wdColorDarkRed = 128,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>32896</remarks>
		 [SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		 wdColorDarkYellow = 32896,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>13209</remarks>
		 [SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		 wdColorBrown = 13209,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>13107</remarks>
		 [SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		 wdColorOliveGreen = 13107,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>13056</remarks>
		 [SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		 wdColorDarkGreen = 13056,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>6697728</remarks>
		 [SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		 wdColorDarkTeal = 6697728,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>10040115</remarks>
		 [SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		 wdColorIndigo = 10040115,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>26367</remarks>
		 [SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		 wdColorOrange = 26367,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>10053222</remarks>
		 [SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		 wdColorBlueGray = 10053222,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>39423</remarks>
		 [SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		 wdColorLightOrange = 39423,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>52377</remarks>
		 [SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		 wdColorLime = 52377,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>6723891</remarks>
		 [SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		 wdColorSeaGreen = 6723891,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>13421619</remarks>
		 [SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		 wdColorAqua = 13421619,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>16737843</remarks>
		 [SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		 wdColorLightBlue = 16737843,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>52479</remarks>
		 [SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		 wdColorGold = 52479,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>16763904</remarks>
		 [SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		 wdColorSkyBlue = 16763904,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>6697881</remarks>
		 [SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		 wdColorPlum = 6697881,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>13408767</remarks>
		 [SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		 wdColorRose = 13408767,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>10079487</remarks>
		 [SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		 wdColorTan = 10079487,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>10092543</remarks>
		 [SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		 wdColorLightYellow = 10092543,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>13434828</remarks>
		 [SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		 wdColorLightGreen = 13434828,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>16777164</remarks>
		 [SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		 wdColorLightTurquoise = 16777164,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>16764057</remarks>
		 [SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		 wdColorPaleBlue = 16764057,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>16751052</remarks>
		 [SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		 wdColorLavender = 16751052,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>15987699</remarks>
		 [SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		 wdColorGray05 = 15987699,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>15132390</remarks>
		 [SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		 wdColorGray10 = 15132390,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>14737632</remarks>
		 [SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		 wdColorGray125 = 14737632,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>14277081</remarks>
		 [SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		 wdColorGray15 = 14277081,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>13421772</remarks>
		 [SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		 wdColorGray20 = 13421772,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>12632256</remarks>
		 [SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		 wdColorGray25 = 12632256,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>11776947</remarks>
		 [SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		 wdColorGray30 = 11776947,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>10921638</remarks>
		 [SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		 wdColorGray35 = 10921638,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>10526880</remarks>
		 [SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		 wdColorGray375 = 10526880,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>10066329</remarks>
		 [SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		 wdColorGray40 = 10066329,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>9211020</remarks>
		 [SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		 wdColorGray45 = 9211020,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>8421504</remarks>
		 [SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		 wdColorGray50 = 8421504,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>7566195</remarks>
		 [SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		 wdColorGray55 = 7566195,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>6710886</remarks>
		 [SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		 wdColorGray60 = 6710886,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>6316128</remarks>
		 [SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		 wdColorGray625 = 6316128,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>5855577</remarks>
		 [SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		 wdColorGray65 = 5855577,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>5000268</remarks>
		 [SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		 wdColorGray70 = 5000268,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>4210752</remarks>
		 [SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		 wdColorGray75 = 4210752,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>3355443</remarks>
		 [SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		 wdColorGray80 = 3355443,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>2500134</remarks>
		 [SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		 wdColorGray85 = 2500134,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>2105376</remarks>
		 [SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		 wdColorGray875 = 2105376,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>1644825</remarks>
		 [SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		 wdColorGray90 = 1644825,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>789516</remarks>
		 [SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		 wdColorGray95 = 789516
	}
}