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
	public enum ChartCommandIdEnum
	{
		 /// <summary>
		 /// SupportByVersion OWC10 1
		 /// </summary>
		 /// <remarks>1001</remarks>
		 [SupportByVersion("OWC10", 1)]
		 chCommandCut = 1001,

		 /// <summary>
		 /// SupportByVersion OWC10 1
		 /// </summary>
		 /// <remarks>1011</remarks>
		 [SupportByVersion("OWC10", 1)]
		 chCommandDeleteSelection = 1011,

		 /// <summary>
		 /// SupportByVersion OWC10 1
		 /// </summary>
		 /// <remarks>1005</remarks>
		 [SupportByVersion("OWC10", 1)]
		 chCommandShowPropertyToolbox = 1005,

		 /// <summary>
		 /// SupportByVersion OWC10 1
		 /// </summary>
		 /// <remarks>6001</remarks>
		 [SupportByVersion("OWC10", 1)]
		 chCommandShowContextMenu = 6001,

		 /// <summary>
		 /// SupportByVersion OWC10 1
		 /// </summary>
		 /// <remarks>1000</remarks>
		 [SupportByVersion("OWC10", 1)]
		 chCommandUndo = 1000,

		 /// <summary>
		 /// SupportByVersion OWC10 1
		 /// </summary>
		 /// <remarks>6002</remarks>
		 [SupportByVersion("OWC10", 1)]
		 chCommandSelectPrevMinor = 6002,

		 /// <summary>
		 /// SupportByVersion OWC10 1
		 /// </summary>
		 /// <remarks>6003</remarks>
		 [SupportByVersion("OWC10", 1)]
		 chCommandSelectNextMinor = 6003,

		 /// <summary>
		 /// SupportByVersion OWC10 1
		 /// </summary>
		 /// <remarks>6004</remarks>
		 [SupportByVersion("OWC10", 1)]
		 chCommandSelectPrevMajor = 6004,

		 /// <summary>
		 /// SupportByVersion OWC10 1
		 /// </summary>
		 /// <remarks>6005</remarks>
		 [SupportByVersion("OWC10", 1)]
		 chCommandSelectNextMajor = 6005,

		 /// <summary>
		 /// SupportByVersion OWC10 1
		 /// </summary>
		 /// <remarks>1006</remarks>
		 [SupportByVersion("OWC10", 1)]
		 chCommandShowHelp = 1006,

		 /// <summary>
		 /// SupportByVersion OWC10 1
		 /// </summary>
		 /// <remarks>1007</remarks>
		 [SupportByVersion("OWC10", 1)]
		 chCommandShowAbout = 1007,

		 /// <summary>
		 /// SupportByVersion OWC10 1
		 /// </summary>
		 /// <remarks>6026</remarks>
		 [SupportByVersion("OWC10", 1)]
		 chCommandPassiveAlert = 6026,

		 /// <summary>
		 /// SupportByVersion OWC10 1
		 /// </summary>
		 /// <remarks>6027</remarks>
		 [SupportByVersion("OWC10", 1)]
		 chCommandLaunchDataFinder = 6027,

		 /// <summary>
		 /// SupportByVersion OWC10 1
		 /// </summary>
		 /// <remarks>6028</remarks>
		 [SupportByVersion("OWC10", 1)]
		 chCommandShowLegend = 6028,

		 /// <summary>
		 /// SupportByVersion OWC10 1
		 /// </summary>
		 /// <remarks>1014</remarks>
		 [SupportByVersion("OWC10", 1)]
		 chCommandRefresh = 1014,

		 /// <summary>
		 /// SupportByVersion OWC10 1
		 /// </summary>
		 /// <remarks>6032</remarks>
		 [SupportByVersion("OWC10", 1)]
		 chCommandByRowCol = 6032,

		 /// <summary>
		 /// SupportByVersion OWC10 1
		 /// </summary>
		 /// <remarks>2000</remarks>
		 [SupportByVersion("OWC10", 1)]
		 chCommandSortAscending = 2000,

		 /// <summary>
		 /// SupportByVersion OWC10 1
		 /// </summary>
		 /// <remarks>2031</remarks>
		 [SupportByVersion("OWC10", 1)]
		 chCommandSortDescending = 2031,

		 /// <summary>
		 /// SupportByVersion OWC10 1
		 /// </summary>
		 /// <remarks>1017</remarks>
		 [SupportByVersion("OWC10", 1)]
		 chCommandAutoFilter = 1017,

		 /// <summary>
		 /// SupportByVersion OWC10 1
		 /// </summary>
		 /// <remarks>1016</remarks>
		 [SupportByVersion("OWC10", 1)]
		 chCommandAutoCalc = 1016,

		 /// <summary>
		 /// SupportByVersion OWC10 1
		 /// </summary>
		 /// <remarks>1012</remarks>
		 [SupportByVersion("OWC10", 1)]
		 chCommandExpand = 1012,

		 /// <summary>
		 /// SupportByVersion OWC10 1
		 /// </summary>
		 /// <remarks>1013</remarks>
		 [SupportByVersion("OWC10", 1)]
		 chCommandCollapse = 1013,

		 /// <summary>
		 /// SupportByVersion OWC10 1
		 /// </summary>
		 /// <remarks>6034</remarks>
		 [SupportByVersion("OWC10", 1)]
		 chCommandDrill = 6034,

		 /// <summary>
		 /// SupportByVersion OWC10 1
		 /// </summary>
		 /// <remarks>1010</remarks>
		 [SupportByVersion("OWC10", 1)]
		 chCommandFieldList = 1010,

		 /// <summary>
		 /// SupportByVersion OWC10 1
		 /// </summary>
		 /// <remarks>1015</remarks>
		 [SupportByVersion("OWC10", 1)]
		 chCommandFilterByMenu = 1015,

		 /// <summary>
		 /// SupportByVersion OWC10 1
		 /// </summary>
		 /// <remarks>6035</remarks>
		 [SupportByVersion("OWC10", 1)]
		 chCommandSortAscendingByTotal = 6035,

		 /// <summary>
		 /// SupportByVersion OWC10 1
		 /// </summary>
		 /// <remarks>6036</remarks>
		 [SupportByVersion("OWC10", 1)]
		 chCommandSortDescendingByTotal = 6036,

		 /// <summary>
		 /// SupportByVersion OWC10 1
		 /// </summary>
		 /// <remarks>6037</remarks>
		 [SupportByVersion("OWC10", 1)]
		 chCommandDrillOut = 6037,

		 /// <summary>
		 /// SupportByVersion OWC10 1
		 /// </summary>
		 /// <remarks>6038</remarks>
		 [SupportByVersion("OWC10", 1)]
		 chCommandTogglePropertiesInScreenTip = 6038,

		 /// <summary>
		 /// SupportByVersion OWC10 1
		 /// </summary>
		 /// <remarks>6039</remarks>
		 [SupportByVersion("OWC10", 1)]
		 chCommandChartType = 6039,

		 /// <summary>
		 /// SupportByVersion OWC10 1
		 /// </summary>
		 /// <remarks>6040</remarks>
		 [SupportByVersion("OWC10", 1)]
		 chCommandShowWizard = 6040,

		 /// <summary>
		 /// SupportByVersion OWC10 1
		 /// </summary>
		 /// <remarks>6041</remarks>
		 [SupportByVersion("OWC10", 1)]
		 chCommandSum = 6041,

		 /// <summary>
		 /// SupportByVersion OWC10 1
		 /// </summary>
		 /// <remarks>6042</remarks>
		 [SupportByVersion("OWC10", 1)]
		 chCommandCount = 6042,

		 /// <summary>
		 /// SupportByVersion OWC10 1
		 /// </summary>
		 /// <remarks>6043</remarks>
		 [SupportByVersion("OWC10", 1)]
		 chCommandMin = 6043,

		 /// <summary>
		 /// SupportByVersion OWC10 1
		 /// </summary>
		 /// <remarks>6044</remarks>
		 [SupportByVersion("OWC10", 1)]
		 chCommandMax = 6044,

		 /// <summary>
		 /// SupportByVersion OWC10 1
		 /// </summary>
		 /// <remarks>6045</remarks>
		 [SupportByVersion("OWC10", 1)]
		 chCommandAverage = 6045,

		 /// <summary>
		 /// SupportByVersion OWC10 1
		 /// </summary>
		 /// <remarks>6046</remarks>
		 [SupportByVersion("OWC10", 1)]
		 chCommandStdDev = 6046,

		 /// <summary>
		 /// SupportByVersion OWC10 1
		 /// </summary>
		 /// <remarks>6047</remarks>
		 [SupportByVersion("OWC10", 1)]
		 chCommandVar = 6047,

		 /// <summary>
		 /// SupportByVersion OWC10 1
		 /// </summary>
		 /// <remarks>6048</remarks>
		 [SupportByVersion("OWC10", 1)]
		 chCommandStdDevP = 6048,

		 /// <summary>
		 /// SupportByVersion OWC10 1
		 /// </summary>
		 /// <remarks>6049</remarks>
		 [SupportByVersion("OWC10", 1)]
		 chCommandVarP = 6049,

		 /// <summary>
		 /// SupportByVersion OWC10 1
		 /// </summary>
		 /// <remarks>1050</remarks>
		 [SupportByVersion("OWC10", 1)]
		 chCommandFontName = 1050,

		 /// <summary>
		 /// SupportByVersion OWC10 1
		 /// </summary>
		 /// <remarks>1051</remarks>
		 [SupportByVersion("OWC10", 1)]
		 chCommandFontSize = 1051,

		 /// <summary>
		 /// SupportByVersion OWC10 1
		 /// </summary>
		 /// <remarks>1052</remarks>
		 [SupportByVersion("OWC10", 1)]
		 chCommandBold = 1052,

		 /// <summary>
		 /// SupportByVersion OWC10 1
		 /// </summary>
		 /// <remarks>1053</remarks>
		 [SupportByVersion("OWC10", 1)]
		 chCommandItalic = 1053,

		 /// <summary>
		 /// SupportByVersion OWC10 1
		 /// </summary>
		 /// <remarks>1054</remarks>
		 [SupportByVersion("OWC10", 1)]
		 chCommandUnderline = 1054,

		 /// <summary>
		 /// SupportByVersion OWC10 1
		 /// </summary>
		 /// <remarks>1055</remarks>
		 [SupportByVersion("OWC10", 1)]
		 chCommandLineColor = 1055,

		 /// <summary>
		 /// SupportByVersion OWC10 1
		 /// </summary>
		 /// <remarks>1056</remarks>
		 [SupportByVersion("OWC10", 1)]
		 chCommandInteriorColor = 1056,

		 /// <summary>
		 /// SupportByVersion OWC10 1
		 /// </summary>
		 /// <remarks>1057</remarks>
		 [SupportByVersion("OWC10", 1)]
		 chCommandFontColor = 1057,

		 /// <summary>
		 /// SupportByVersion OWC10 1
		 /// </summary>
		 /// <remarks>6050</remarks>
		 [SupportByVersion("OWC10", 1)]
		 chCommandMultiChart = 6050,

		 /// <summary>
		 /// SupportByVersion OWC10 1
		 /// </summary>
		 /// <remarks>6051</remarks>
		 [SupportByVersion("OWC10", 1)]
		 chCommandUnifiedScales = 6051,

		 /// <summary>
		 /// SupportByVersion OWC10 1
		 /// </summary>
		 /// <remarks>6052</remarks>
		 [SupportByVersion("OWC10", 1)]
		 chCommandShowDropZones = 6052,

		 /// <summary>
		 /// SupportByVersion OWC10 1
		 /// </summary>
		 /// <remarks>6053</remarks>
		 [SupportByVersion("OWC10", 1)]
		 chCommandShowToolbar = 6053,

		 /// <summary>
		 /// SupportByVersion OWC10 1
		 /// </summary>
		 /// <remarks>1100</remarks>
		 [SupportByVersion("OWC10", 1)]
		 chCommandShowTop1 = 1100,

		 /// <summary>
		 /// SupportByVersion OWC10 1
		 /// </summary>
		 /// <remarks>1101</remarks>
		 [SupportByVersion("OWC10", 1)]
		 chCommandShowTop2 = 1101,

		 /// <summary>
		 /// SupportByVersion OWC10 1
		 /// </summary>
		 /// <remarks>1102</remarks>
		 [SupportByVersion("OWC10", 1)]
		 chCommandShowTop5 = 1102,

		 /// <summary>
		 /// SupportByVersion OWC10 1
		 /// </summary>
		 /// <remarks>1103</remarks>
		 [SupportByVersion("OWC10", 1)]
		 chCommandShowTop10 = 1103,

		 /// <summary>
		 /// SupportByVersion OWC10 1
		 /// </summary>
		 /// <remarks>1104</remarks>
		 [SupportByVersion("OWC10", 1)]
		 chCommandShowTop25 = 1104,

		 /// <summary>
		 /// SupportByVersion OWC10 1
		 /// </summary>
		 /// <remarks>1105</remarks>
		 [SupportByVersion("OWC10", 1)]
		 chCommandShowTop1Percent = 1105,

		 /// <summary>
		 /// SupportByVersion OWC10 1
		 /// </summary>
		 /// <remarks>1106</remarks>
		 [SupportByVersion("OWC10", 1)]
		 chCommandShowTop2Percent = 1106,

		 /// <summary>
		 /// SupportByVersion OWC10 1
		 /// </summary>
		 /// <remarks>1107</remarks>
		 [SupportByVersion("OWC10", 1)]
		 chCommandShowTop5Percent = 1107,

		 /// <summary>
		 /// SupportByVersion OWC10 1
		 /// </summary>
		 /// <remarks>1108</remarks>
		 [SupportByVersion("OWC10", 1)]
		 chCommandShowTop10Percent = 1108,

		 /// <summary>
		 /// SupportByVersion OWC10 1
		 /// </summary>
		 /// <remarks>1109</remarks>
		 [SupportByVersion("OWC10", 1)]
		 chCommandShowTop25Percent = 1109,

		 /// <summary>
		 /// SupportByVersion OWC10 1
		 /// </summary>
		 /// <remarks>1110</remarks>
		 [SupportByVersion("OWC10", 1)]
		 chCommandShowBottom1 = 1110,

		 /// <summary>
		 /// SupportByVersion OWC10 1
		 /// </summary>
		 /// <remarks>1111</remarks>
		 [SupportByVersion("OWC10", 1)]
		 chCommandShowBottom2 = 1111,

		 /// <summary>
		 /// SupportByVersion OWC10 1
		 /// </summary>
		 /// <remarks>1112</remarks>
		 [SupportByVersion("OWC10", 1)]
		 chCommandShowBottom5 = 1112,

		 /// <summary>
		 /// SupportByVersion OWC10 1
		 /// </summary>
		 /// <remarks>1113</remarks>
		 [SupportByVersion("OWC10", 1)]
		 chCommandShowBottom10 = 1113,

		 /// <summary>
		 /// SupportByVersion OWC10 1
		 /// </summary>
		 /// <remarks>1114</remarks>
		 [SupportByVersion("OWC10", 1)]
		 chCommandShowBottom25 = 1114,

		 /// <summary>
		 /// SupportByVersion OWC10 1
		 /// </summary>
		 /// <remarks>1115</remarks>
		 [SupportByVersion("OWC10", 1)]
		 chCommandShowBottom1Percent = 1115,

		 /// <summary>
		 /// SupportByVersion OWC10 1
		 /// </summary>
		 /// <remarks>1116</remarks>
		 [SupportByVersion("OWC10", 1)]
		 chCommandShowBottom2Percent = 1116,

		 /// <summary>
		 /// SupportByVersion OWC10 1
		 /// </summary>
		 /// <remarks>1117</remarks>
		 [SupportByVersion("OWC10", 1)]
		 chCommandShowBottom5Percent = 1117,

		 /// <summary>
		 /// SupportByVersion OWC10 1
		 /// </summary>
		 /// <remarks>1118</remarks>
		 [SupportByVersion("OWC10", 1)]
		 chCommandShowBottom10Percent = 1118,

		 /// <summary>
		 /// SupportByVersion OWC10 1
		 /// </summary>
		 /// <remarks>1119</remarks>
		 [SupportByVersion("OWC10", 1)]
		 chCommandShowBottom25Percent = 1119,

		 /// <summary>
		 /// SupportByVersion OWC10 1
		 /// </summary>
		 /// <remarks>1120</remarks>
		 [SupportByVersion("OWC10", 1)]
		 chCommandShowOther = 1120,

		 /// <summary>
		 /// SupportByVersion OWC10 1
		 /// </summary>
		 /// <remarks>1121</remarks>
		 [SupportByVersion("OWC10", 1)]
		 chCommandShowAll = 1121,

		 /// <summary>
		 /// SupportByVersion OWC10 1
		 /// </summary>
		 /// <remarks>1123</remarks>
		 [SupportByVersion("OWC10", 1)]
		 chCommandShowTopNMenu = 1123,

		 /// <summary>
		 /// SupportByVersion OWC10 1
		 /// </summary>
		 /// <remarks>1124</remarks>
		 [SupportByVersion("OWC10", 1)]
		 chCommandShowBottomNMenu = 1124,

		 /// <summary>
		 /// SupportByVersion OWC10 1
		 /// </summary>
		 /// <remarks>1125</remarks>
		 [SupportByVersion("OWC10", 1)]
		 chCommandConditionalFilter = 1125,

		 /// <summary>
		 /// SupportByVersion OWC10 1
		 /// </summary>
		 /// <remarks>6054</remarks>
		 [SupportByVersion("OWC10", 1)]
		 chCommandMoveToFilterArea = 6054,

		 /// <summary>
		 /// SupportByVersion OWC10 1
		 /// </summary>
		 /// <remarks>6055</remarks>
		 [SupportByVersion("OWC10", 1)]
		 chCommandMoveToSeriesArea = 6055,

		 /// <summary>
		 /// SupportByVersion OWC10 1
		 /// </summary>
		 /// <remarks>6056</remarks>
		 [SupportByVersion("OWC10", 1)]
		 chCommandMoveToCategoryArea = 6056,

		 /// <summary>
		 /// SupportByVersion OWC10 1
		 /// </summary>
		 /// <remarks>6057</remarks>
		 [SupportByVersion("OWC10", 1)]
		 chCommandMoveToChartArea = 6057
	}
}